"""Portable source/archive preparation when provider-bundled helpers are unavailable.

Credentials arrive on hidden stdin, remain in process memory, and are passed to
Git through per-process configuration environment. No token is written to disk.
Only the owner-private preview checkout is operated on; no force pushes.
"""
import json
import os
import subprocess
import sys
import termios
from pathlib import Path

checkout = Path('/workspace/budgetbook/preview-site')
archive = Path('/workspace/budgetbook/playground/private-preview.tar.gz')
if sys.stdin.isatty():
    options = termios.tcgetattr(sys.stdin.fileno())
    options[3] &= ~termios.ECHO
    termios.tcsetattr(sys.stdin.fileno(), termios.TCSANOW, options)
print('Ready for source credential JSON on stdin (input is hidden).', flush=True)
payload = json.loads(sys.stdin.readline())
credential = payload['credential']
identity = payload['project_id']
manifest = json.loads((checkout / '.openai/hosting.json').read_text())
assert manifest['project_id'] == identity
assert credential['auth_mode'] == 'http_extra_header'
assert credential['remote_url'].startswith('https://')
assert '@' not in credential['remote_url'].split('://', 1)[1]
token = credential['token']
environment = os.environ.copy()
for key in tuple(environment):
    if key.startswith('GIT_TRACE') or key in ('GIT_CURL_VERBOSE', 'GIT_CONFIG_COUNT'):
        environment.pop(key, None)
environment.update({
    'GIT_CONFIG_COUNT': '1',
    'GIT_CONFIG_KEY_0': 'http.extraHeader',
    'GIT_CONFIG_VALUE_0': 'Authorization: Bearer ' + token,
    'GIT_TERMINAL_PROMPT': '0',
    'GIT_AUTHOR_NAME': 'Codex',
    'GIT_AUTHOR_EMAIL': 'codex@local.invalid',
    'GIT_COMMITTER_NAME': 'Codex',
    'GIT_COMMITTER_EMAIL': 'codex@local.invalid',
})

def git(*arguments):
    completed = subprocess.run(['git', *arguments], cwd=checkout,
                               env=environment, text=True, capture_output=True)
    if completed.returncode:
        message = (completed.stderr + completed.stdout).replace(token, '[redacted]')
        raise RuntimeError(message[:2500])
    return completed.stdout.strip()

if not (checkout / '.git').exists():
    git('init', '--initial-branch=' + credential['branch'])
remotes = git('remote').splitlines()
if 'origin' in remotes:
    assert git('remote', 'get-url', 'origin') == credential['remote_url']
else:
    git('remote', 'add', 'origin', credential['remote_url'])
if payload.get('mode') == 'open':
    sha = git('rev-parse', 'HEAD')
    remote_head = git('ls-remote', 'origin', 'refs/heads/' + credential['branch'])
    assert remote_head.split()[0] == sha, 'Remote source changed; synchronize before editing.'
    assert not git('status', '--porcelain', '--untracked-files=no')
    print(json.dumps({'project_id': identity, 'checkout_path': str(checkout), 'commit_sha': sha}))
    sys.exit(0)
git('add', '.openai/hosting.json', 'dist/index.html')
if git('diff', '--cached', '--name-only'):
    git('commit', '-m', 'Add owner-private functional budget preview')
sha = git('rev-parse', 'HEAD')
git('push', 'origin', 'HEAD:refs/heads/' + credential['branch'])
remote_head = git('ls-remote', 'origin', 'refs/heads/' + credential['branch'])
assert remote_head.split()[0] == sha
assert not git('status', '--porcelain', '--untracked-files=no')
git('archive', '--format=tar.gz', '--output=' + str(archive), sha,
    '.openai/hosting.json', 'dist/index.html')
print(json.dumps({'project_id': identity, 'checkout_path': str(checkout),
                  'commit_sha': sha, 'archive': str(archive)}, ensure_ascii=False))
