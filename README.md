# QueueSwitcher

##Features
select number of buttons to display

##Limitations
Binds to a field that isnt used.
Does not auto refresh when the form refreshes if "add to queue" button is used
No message if not attached to a queue
might possibly need to include inactive queue items when getting top 50 recent records
not currently tidying up using the destroy() function

## Enable To
Bind to any single text field as it isn't used
- SingleLine.Text

##Setup
Use the solution.zip to install managed component to your environment

##Solution creation
navigate to QueueSwitcherSolution folder
pac solution init --publisher-name XXX --publisher-prefix yyy

pac solution add-reference --path ../

msbuild /t:build /restore

##Solution install
Manually download zip file from releases and import.

##Solution Deployment
To deploy to your environment using the command line
copy the repository locally
connect to remote env
pac auth create --url XXX
enter credentials
command to push to env
pac pcf push --publisher-prefix lsl
