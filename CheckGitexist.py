import requests

def check_file_exists(repo_owner, repo_name, path, token):
    url = f"https://api.github.com/repos/{repo_owner}/{repo_name}/contents/{path}"
    headers = {"Authorization": f"token {token}"}
    response = requests.get(url, headers=headers)

    if response.status_code == 200:
        print(f"The file '{path}' exists in the repository '{repo_name}'.")
    else:
        print(f"The file '{path}' does not exist in the repository '{repo_name}'. Response code: {response.status_code}")

# Replace these variables with your specific details
repo_owner = "ocean-network-express"
repo_name = "LOOKML_one_vessel_special_cargo_spoke"
path = "manifest_lock.lkml"
token = "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"

check_file_exists(repo_owner, repo_name, path, token)
