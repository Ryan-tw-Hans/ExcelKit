from concurrent.futures import ThreadPoolExecutor
from pathlib import Path
import subprocess
import tempfile

root = Path(__file__).resolve().parents[2]
subprocess.run(['powershell.exe', '-NoProfile', '-NonInteractive', '-File', str(root / 'Installer/Tests/Test-GitHubPackageVersion.ps1')], check=True)

with tempfile.TemporaryDirectory(prefix='excelkit-ref-race-') as tmp:
    tmp = Path(tmp)
    remote = tmp / 'remote.git'
    subprocess.run(['git', 'init', '--bare', str(remote)], check=True, capture_output=True)
    sources = []
    for name in ('a', 'b'):
        path = tmp / name
        subprocess.run(['git', 'init', str(path)], check=True, capture_output=True)
        subprocess.run(['git', '-C', str(path), 'config', 'user.name', 'Test'], check=True)
        subprocess.run(['git', '-C', str(path), 'config', 'user.email', 'test@example.invalid'], check=True)
        (path / 'input.txt').write_text(name, encoding='utf-8')
        subprocess.run(['git', '-C', str(path), 'add', 'input.txt'], check=True)
        subprocess.run(['git', '-C', str(path), 'commit', '-m', 'fixture'], check=True, capture_output=True)
        sources.append(path)
    def push(path):
        return subprocess.run(['git', '-C', str(path), 'push', str(remote), 'HEAD:refs/tags/excelkit-package/20990101/01'], capture_output=True).returncode
    with ThreadPoolExecutor(max_workers=2) as executor:
        results = list(executor.map(push, sources))
    assert sum(code == 0 for code in results) == 1, results
    print('PASS concurrent Git ref creation permits exactly one writer')
