"""GitHub workflow helper: prints the app version and, for a tag build, checks that the tag matches it.
Pushing tag v2.3.3 needs VERSION = '2.3.3' in mouse_colony_tool.py, so an installer never carries a wrong number."""
import os
import re
import sys

src = open(os.path.join(os.path.dirname(__file__), '..', 'mouse_colony_tool.py'), encoding='utf-8').read()
version = re.search(r"^VERSION = '([^']+)'", src, re.M).group(1)
if os.environ.get('GITHUB_REF_TYPE') == 'tag':
    tag = os.environ.get('GITHUB_REF_NAME', '').lstrip('vV')
    if tag != version:
        print(f"::error::Tag v{tag} does not match VERSION = '{version}' in mouse_colony_tool.py. "
              f"Change one of them so they match, then push the tag again.")
        sys.exit(1)
print(f'Version {version}')
out = os.environ.get('GITHUB_OUTPUT')
if out:
    with open(out, 'a', encoding='utf-8') as fh:
        fh.write(f'version={version}\n')
