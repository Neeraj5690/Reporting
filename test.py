from github import Github

g = Github("Neeraj5690", "Neeraj@408")
g = Github("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")

#  All repos present
for repo in g.get_user().get_repos():
    print(repo.name)

try:
    # Accessing particular Repo and its folder
    repo=g.get_repo("Neeraj5690/Reporting")
    # Removing files from the folder
    Folder=repo.get_contents("/ReportData")
    for contentFiles in Folder:
        print(contentFiles)

        repo.delete_file(contentFiles.path, "message", contentFiles.sha, branch='master')
except:
    print("No File present")