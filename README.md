# QueueSwitcher

##Features

##Limitations

## Enable To
Hoping to bind to anything as it isnt used
- SingleLine.Text

##Setup
Use the solution.zip to install managed component to your environment

##Solution creation
navigate to QueueSwitcherSolution folder
pac solution init --publisher-name LusciniaSolutionsLimited --publisher-prefix lsl

pac solution add-reference --path ../

msbuild /t:build /restore

##Solution install


##Solution Deployment
To deploy to your environment using the command line
copy the repository locally
connect to remote env
pac auth create --url https://lusciniadev.crm11.dynamics.com/
enter credentials
command to push to env
pac pcf push --publisher-prefix lsl