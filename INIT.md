# Github and Pypi

## Create Pypi Pending Publisher

https://pypi.org/manage/account/publishing/
* Project Name: openpyxl-table
* Owner: KalleDK
* Repository name: py-openpyxl-table
* Workflow name: release.yml
* Environment name: pypi

https://test.pypi.org/manage/account/publishing/
* Project Name: openpyxl-table
* Owner: KalleDK
* Repository name: py-openpyxl-table
* Workflow name: testrelease.yml
* Environment name: testpypi

## Create Github Repo
```sh
gh repo create py-openpyxl-table --public --source=. --remote=upstream
```

## Create Github Environment for Pypi

```sh
echo '
import subprocess
ENV_NAME="pypi"

subprocess.run([
  "gh", "api",
  "-X", "PUT",
  "-H", "Accept: application/vnd.github+json",
  f"/repos/:owner/:repo/environments/{ENV_NAME}",
  "--input", "-"
],
input=b''{"deployment_branch_policy": {"protected_branches": false, "custom_branch_policies": true}}''
)

subprocess.run([
  "gh", "api",
  "-X", "POST",
  "-H", "Accept: application/vnd.github+json",
  f"/repos/:owner/:repo/environments/{ENV_NAME}/deployment-branch-policies",
  "-f", "name=v*",
  "-f", "type=tag"
])
' | uv run python
```

## Create Github Environment for Pypi Test

```sh
echo '
import subprocess
ENV_NAME="testpypi"

subprocess.run([
  "gh", "api",
  "-X", "PUT",
  "-H", "Accept: application/vnd.github+json",
  f"/repos/:owner/:repo/environments/{ENV_NAME}",
  "--input", "-"
],
input=b''{"deployment_branch_policy": {"protected_branches": false, "custom_branch_policies": true}}''
)

subprocess.run([
  "gh", "api",
  "-X", "POST",
  "-H", "Accept: application/vnd.github+json",
  f"/repos/:owner/:repo/environments/{ENV_NAME}/deployment-branch-policies",
  "-f", "name=v*",
  "-f", "type=tag"
])
' | uv run python
```
