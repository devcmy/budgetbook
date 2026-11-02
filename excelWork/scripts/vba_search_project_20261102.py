"""Add one source-only search module, preserving every existing VBA source.

No macro is run here. MS-CFB output is independently parsed with olefile and
MS-OVBA metadata/source are independently parsed with oletools before delivery.
Only the reconstructible VBA performance caches are discarded.
Dependencies for build/verification: olefile, oletools (not needed in Excel).
"""
from io import BytesIO
from math import ceil
import struct
import uuid

import olefile
from oletools.olevba import VBA_Project, VBA_Module, decompress_stream

FREE=0xFFFFFFFF
END=0xFFFFFFFE
FAT=0xFFFFFFFD


def compress(data):
    """MS-OVBA compressed container, with greedy LZ tokens and raw full chunks."""
    out=bytearray(b'\x01')
    for start in range(0,len(data),4096):
        block=data[start:start+4096];payload=bytearray();pos=0
        while pos<len(block):
            flagpos=len(payload);payload.append(0);flags=0
            for bit in range(8):
                if pos>=len(block):break
                offsetbits=max(4,(pos-1).bit_length())
                lengthbits=16-offsetbits;limit=min((1<<lengthbits)+2,len(block)-pos)
                best=0;distance=0
                if limit>=3:
                    needle=block[pos:pos+3];at=block.rfind(needle,0,pos)
                    while at>=0:
                        size=3
                        while size<limit and block[at+size]==block[pos+size]:size+=1
                        if size>best:best=size;distance=pos-at
                        if best==limit:break
                        at=block.rfind(needle,0,at)
                if best>=3:
                    flags|=1<<bit
                    payload.extend(struct.pack('<H',((distance-1)<<lengthbits)|(best-3)))
                    pos+=best
                else:payload.append(block[pos]);pos+=1
            payload[flagpos]=flags
        if len(payload)>=4096 and len(block)==4096:
            out.extend(struct.pack('<H',0x3FFF));out.extend(block)
        else:
            assert len(payload)<=4096,'Incompressible short VBA chunk'
            out.extend(struct.pack('<H',0xB000|(len(payload)-1)));out.extend(payload)
    assert decompress_stream(out)==data
    return bytes(out)


def compound(streams, storage_metadata=None):
    """Small MS-CFB v3 writer, including mini streams and red/black directories."""
    storage_metadata=storage_metadata or {}
    names={():{'type':5,'children':[]}}
    for path,data in streams.items():
        for depth in range(1,len(path)):
            names.setdefault(path[:depth],{'type':1,'children':[]})
        names[path]={'type':2,'data':data,'children':[]}
    paths=[()]+sorted((p for p in names if p),key=lambda p:(len(p),p))
    for sid,path in enumerate(paths):
        names[path].update(sid=sid,left=FREE,right=FREE,child=FREE,color=1)
        if path:names[path[:-1]]['children'].append(path)
    for parent in paths:
        kids=sorted(names[parent]['children'],key=lambda p:(len(p[-1]),p[-1].upper()))
        maxdepth=(len(kids)).bit_length()-1
        def tree(items,depth=0):
            if not items:return FREE
            mid=len(items)//2;node=names[items[mid]]
            node['left']=tree(items[:mid],depth+1);node['right']=tree(items[mid+1:],depth+1)
            node['color']=0 if depth==maxdepth and depth>0 else 1
            return node['sid']
        names[parent]['child']=tree(kids)
    sectors=[];links=[]
    def allocate(data):
        if not data:return END
        first=len(sectors);count=ceil(len(data)/512)
        for i in range(count):
            sectors.append(data[i*512:(i+1)*512].ljust(512,b'\x00'))
            links.append(first+i+1 if i<count-1 else END)
        return first
    mini=bytearray();minifat=[]
    for path in paths:
        n=names[path]
        if n['type']!=2:continue
        data=n['data'];n['size']=len(data)
        if not data:n['start']=END
        elif len(data)<4096:
            first=len(minifat);count=ceil(len(data)/64);n['start']=first
            mini.extend(data.ljust(count*64,b'\x00'))
            minifat.extend([first+i+1 if i<count-1 else END for i in range(count)])
        else:n['start']=allocate(data)
    names[()]['start']=allocate(mini);names[()]['size']=len(mini)
    mini_bytes=struct.pack('<'+'I'*len(minifat),*minifat)
    mini_start=allocate(mini_bytes.ljust(ceil(len(mini_bytes)/512)*512,b'\xff'))
    directory=bytearray()
    for path in paths:
        n=names[path];name=path[-1] if path else 'Root Entry'
        encoded=(name+'\x00').encode('utf-16le');assert len(encoded)<=64
        meta=storage_metadata.get(path,{})
        clsid=uuid.UUID(meta.get('clsid') or '00000000-0000-0000-0000-000000000000').bytes_le
        directory.extend(encoded.ljust(64,b'\x00'))
        directory.extend(struct.pack('<HBBIII',len(encoded),n['type'],n['color'],n['left'],n['right'],n['child']))
        directory.extend(clsid)
        directory.extend(struct.pack('<IQQIQ',meta.get('state',0),meta.get('create',0),meta.get('modify',0),n.get('start',END),n.get('size',0)))
    directory_start=allocate(directory)
    nfat=ceil(len(sectors)/128)
    while ceil((len(sectors)+nfat)/128)!=nfat:nfat=ceil((len(sectors)+nfat)/128)
    assert nfat<=109,'DIFAT extension not needed for this small project'
    fat_start=len(sectors);links.extend([FAT]*nfat)
    fatdata=struct.pack('<'+'I'*(nfat*128),*(links+[FREE]*(nfat*128-len(links))))
    sectors.extend(fatdata[i*512:(i+1)*512] for i in range(nfat))
    header=bytearray(b'\xd0\xcf\x11\xe0\xa1\xb1\x1a\xe1'+b'\x00'*16)
    header.extend(struct.pack('<HHHHH',0x003E,3,0xFFFE,9,6));header.extend(b'\x00'*6)
    header.extend(struct.pack('<IIIIIIIII',0,nfat,directory_start,0,4096,mini_start,ceil(len(mini_bytes)/512),END,0))
    header.extend(struct.pack('<109I',*(list(range(fat_start,fat_start+nfat))+[FREE]*(109-nfat))))
    assert len(header)==512
    result=bytes(header)+b''.join(sectors)
    verify=olefile.OleFileIO(BytesIO(result),raise_defects=olefile.DEFECT_INCORRECT)
    assert {tuple(p):verify.openstream(p).read() for p in verify.listdir()}==streams
    verify.close();return result


def rec(kind,value=b''):
    return struct.pack('<HI',kind,len(value))+value


def add_search(original,source):
    ole=olefile.OleFileIO(BytesIO(original))
    streams={tuple(p):ole.openstream(p).read() for p in ole.listdir()}
    project=VBA_Project(ole,'','PROJECT','VBA/dir',relaxed=False)
    project.parse_project_stream();raw=project.dir_stream.getvalue();sizepos=project.dir_stream.tell()
    stream=project.dir_stream
    assert struct.unpack('<I',stream.read(4))[0]==2
    count=struct.unpack('<H',stream.read(2))[0];assert stream.read(6)==rec(0x13,b'\x00\x00')[:6]
    stream.read(2);prefix=bytearray(raw[:stream.tell()]);struct.pack_into('<H',prefix,sizepos+4,count+1)
    records=[];before_sources={}
    for i in range(count):
        start=stream.tell();m=VBA_Module(project,stream,i);end=stream.tell()
        assert m.name!='BBSearch'
        module=bytearray(raw[start:end]);needle=rec(0x31,b'\x00'*4)[:6]
        assert module.count(needle)==1
        offset=module.index(needle)+6;module[offset:offset+4]=b'\x00'*4;records.append(bytes(module))
        before_sources[m.name]=m.code_raw
        streams['VBA',m.streamname]=streams['VBA',m.streamname][m.textoffset:]
    name=b'BBSearch';uni='BBSearch'.encode('utf-16le')
    new=(rec(0x19,name)+rec(0x47,uni)+rec(0x1A,name)+rec(0x32,uni)+rec(0x1C)+rec(0x48)
         +rec(0x31,struct.pack('<I',0))+rec(0x1E,struct.pack('<I',0))+rec(0x2C,b'\xff\xff')+rec(0x21)+rec(0x2B))
    streams['VBA','dir']=compress(bytes(prefix)+b''.join(records)+new+raw[stream.tell():])
    streams['VBA','BBSearch']=compress(source)
    streams['VBA','_VBA_PROJECT']=b'\xcc\x61\xff\xff\x00\x00\x00'
    for path in list(streams):
        if path[-1].startswith('__SRP_'):del streams[path]
    text=streams['PROJECT',].decode(project.codec)
    assert 'Module=BBSearch' not in text
    text=text.replace('Name="VBAProject"','Module=BBSearch\r\nName="VBAProject"')
    text+='BBSearch=0, 0, 0, 0, C\r\n';streams['PROJECT',]=text.encode(project.codec)
    wm=streams['PROJECTwm',];assert wm.endswith(b'\x00\x00')
    streams['PROJECTwm',]=wm[:-2]+name+b'\x00'+uni+b'\x00\x00'+b'\x00\x00'
    metadata={}
    for entry in ole.direntries:
        if entry is not None and entry.entry_type in [1,5]:
            path=() if entry.entry_type==5 else (entry.name,)
            metadata[path]={'clsid':entry.clsid,'state':entry.dwUserFlags,'create':entry.createTime,'modify':entry.modifyTime}
    output=compound(streams,metadata);ole.close()
    check=olefile.OleFileIO(BytesIO(output));p=VBA_Project(check,'','PROJECT','VBA/dir',relaxed=False)
    p.parse_project_stream();list(p.parse_modules());after={m.name:m.code_raw for m in p.modules}
    assert all(after[k]==v for k,v in before_sources.items())
    assert after['BBSearch']==source
    check.close()
    return output,{'existing_vba_source_modules_preserved':len(before_sources),'new_module':'BBSearch',
                   'source_only_rebuild_discards_performance_caches':True,'macro_executed':False}
