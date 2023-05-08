from github import Github

g = Github("Neeraj5690", "Neeraj@408")
g = Github("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")

#  All repos present
for repo in g.get_user().get_repos():
    print(repo.name)

# Accessing particular Repo and its folder
repo=g.get_repo("Neeraj5690/Reporting")
# Removing files from the folder
Folder=repo.get_contents("/ReportData")
for contentFiles in Folder:
    print(contentFiles)
    if contentFiles.path.format() == "ReportData/img.png":
        pass
    else:
        repo.delete_file(contentFiles.path, "message", contentFiles.sha, branch='master')
        print("Deleted file "+contentFiles.path)