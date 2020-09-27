# Introduction 
DimCry's Garage full of PowerShell scripts :)

## Installs
 - Git Windows Client: https://git-scm.com/ 
 - Git Extensions: https://sourceforge.net/projects/gitextensions/

## Clone project
 - Navigate to the directory where you want to clone the respository.
 - Make sure this is a short path, something like <DriveLetter>:\GitHubRepos, because putting it in the default %USERPROFILE%\source\repos can lead to odd errors during build which stem from maximum length.
 - Run `git clone --recursive https://github.com/dimcry/DimCry-Garage` 

## Contribute
 - Open project with Visual Studio Code or with GitHub Desktop
 - Ensure you are on the `master` branch
 - Do a `git pull` before starting your work. Doing this, you`ll have the latest updates downloaded locally on your machine
 - Create a new branch and checkout to it. The recommendation for the branch name should be `u/<Name>/<BriefDescriptionOfChange>`. Eg.: `u/DimCry/InitialRepositoryStructure`
 - Do your changes
 - When done, push the changes to the cloud repository. First you have to stage the changes, after that commit them to the local repository and after that you have to push (or sync) them to the cloud repository
 - When this is done, the new branch `u/<Name>/<BriefDescriptionOfChange>` will be created in the GitHub repository
 - To push the change to the `master` branch, you have to create a `Pull request` from your new created `u/<Name>/<BriefDescriptionOfChange>` branch to the `master` branch
 - At least 1 reviewer is required to approve the changes into the `master` branch

## Useful links
- [How to create ReadMe files](https://docs.microsoft.com/en-us/azure/devops/repos/git/create-a-readme?view=azure-devops)
- [ASP.NET Core](https://github.com/aspnet/Home)
- [Visual Studio Code](https://github.com/Microsoft/vscode)
- [Chakra Core](https://github.com/Microsoft/ChakraCore)