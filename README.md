sfdx-cdillc-plugin
====================

Various sfdx tools by CDI LLC

## Installation
```
sfdx plugins:install sfdx-cdillc-plugin
```

<!-- commands -->
* [`sfdx cdillc:permissions:export [-f <string>] [-p <array>] [-s <array>] [-g <array>] [-i <array>] [-u <string>] [--apiversion <string>] [--verbose] [--json] [--loglevel trace|debug|info|warn|error|fatal|TRACE|DEBUG|INFO|WARN|ERROR|FATAL]`](#sfdx-cdillcpermissionsexport--f-string--p-array--s-array--g-array--i-array--u-string---apiversion-string---verbose---json---loglevel-tracedebuginfowarnerrorfataltracedebuginfowarnerrorfatal)

## `sfdx cdillc:permissions:export [-f <string>] [-p <array>] [-s <array>] [-g <array>] [-i <array>] [-u <string>] [--apiversion <string>] [--verbose] [--json] [--loglevel trace|debug|info|warn|error|fatal|TRACE|DEBUG|INFO|WARN|ERROR|FATAL]`

export excel spreadsheet of profiles and permission sets

```
USAGE
  $ sfdx cdillc:permissions:export [-f <string>] [-p <array>] [-s <array>] [-g <array>] [-i <array>] [-u <string>] 
  [--apiversion <string>] [--verbose] [--json] [--loglevel 
  trace|debug|info|warn|error|fatal|TRACE|DEBUG|INFO|WARN|ERROR|FATAL]

OPTIONS
  -f, --filepath=filepath
      [default: permissions.xlsx] file path for the output file. defaults to current directory. the directory path must
      exist.

  -g, --permissionsetgroupnames=permissionsetgroupnames
      a comma delimited list of permission set group names. enclose in quotes if permission set group names contain
      spaces.

  -i, --includedcomponents=includedcomponents
      [default: all] a comma delimited list of profile or permission set components to include in output file. all types
      are included if this option is omitted. valid types include the following: applicationVisibilities, classAccesses,
      customMetadataTypeAccesses, customPermissions, customSettingAccesses, fieldPermissions, flowAccesses,
      layoutAssignments, objectPermissions, pageAccesses, recordTypeVisibilities, tabSettings, userPermissions

  -p, --profilenames=profilenames
      a comma delimited list of profile names. enclose in quotes if profile names contain spaces.

  -s, --permissionsetnames=permissionsetnames
      a comma delimited list of permission set names. enclose in quotes if permission set names contain spaces.

  -u, --targetusername=targetusername
      username or alias for the target org; overrides default target org

  --apiversion=apiversion
      override the api version used for api requests made by this command

  --json
      format output as json

  --loglevel=(trace|debug|info|warn|error|fatal|TRACE|DEBUG|INFO|WARN|ERROR|FATAL)
      [default: warn] logging level for this command invocation

  --verbose
      emit additional command output to stdout

EXAMPLES
  sfdx cdillc:permissions:export -s PS1,PS2,PS3 -p P1,P2 -u myorg@example.com
  sfdx cdillc:permissions:export -s PS1,PS2,PS3 -p P1,P2 -f ./output/permissionsfilename.xlsx -u myorg@example.com
```

_See code: [src/commands/cdillc/permissions/export.ts](https://github.com/Candoris/sfdx-cdillc-plugin/blob/v1.2.0/src/commands/cdillc/permissions/export.ts)_
<!-- commandsstop -->
<!-- debugging-your-plugin -->
# Debugging your plugin
We recommend using the Visual Studio Code (VS Code) IDE for your plugin development. Included in the `.vscode` directory of this plugin is a `launch.json` config file, which allows you to attach a debugger to the node process when running your commands.

To debug the `hello:org` command: 
1. Start the inspector
  
If you linked your plugin to the sfdx cli, call your command with the `dev-suspend` switch: 
```sh-session
$ sfdx hello:org -u myOrg@example.com --dev-suspend
```
  
Alternatively, to call your command using the `bin/run` script, set the `NODE_OPTIONS` environment variable to `--inspect-brk` when starting the debugger:
```sh-session
$ NODE_OPTIONS=--inspect-brk bin/run hello:org -u myOrg@example.com
```

2. Set some breakpoints in your command code
3. Click on the Debug icon in the Activity Bar on the side of VS Code to open up the Debug view.
4. In the upper left hand corner of VS Code, verify that the "Attach to Remote" launch configuration has been chosen.
5. Hit the green play button to the left of the "Attach to Remote" launch configuration window. The debugger should now be suspended on the first line of the program. 
6. Hit the green play button at the top middle of VS Code (this play button will be to the right of the play button that you clicked in step #5).
<br><img src=".images/vscodeScreenshot.png" width="480" height="278"><br>
Congrats, you are debugging!
