##  Result Export Tool

A desktop tool to export a single test result into different formats.
* `excel`: an Excel file with all the screenshots and parameters
* `html`: a zip file with all the page HTML files
* `diag`: a zip file with all evidences and page knowledges for diagnosis

### Build

`make.sh`

### Usage

`java -jar ResultExport.jar <path to config file> <resultFormat> <path to target file>`

* `resultFormat`: type of the result format, `excel`, `html`, or `diag`

**Hint**: if the size of the test result is quite large, please adjust the JVM arguments to increase the heap size. For instance:

`java -Xms1024m -Xmx1024m -jar ResultExport.jar <path to config file> <resultFormat> <path to target file>`

### Config file

#### Parameters

* `serverUrl`: the URL of SWATHub Server URL, such as http://www.swathub.com/
* `username`: the username of SWATHub Server
* `apiKey`: the api key for the user, same as the key for execution node
* `workspaceOwner`: the owner's username of the target workspace to export
* `workspaceName`: the name of the target workspace
* `locale`: the locale to fetch the test result, supporting `en`, `ja` and `zh_cn`

#### Sample config file

```
{
  "serverUrl": "http://swathub.com/",
  "username": "tester",
  "apiKey": "A7185B82DB6A4EFC9006",
  "workspaceOwner": "support",
  "workspaceName": "samples",
  "locale": "en"
}
```

### Target file

#### Parameters

* `ids`: the list of the unique ID of the test result. For instance, the result ID is **12345** in the url `http://swathub.com/app/support/samples/results/12345`.
_Note: the following parameters will not be affected when ids is not blank._
* `setID`: the unique ID (string) of the test set in the target workspace, which can be got from the test set url. For instance, the set ID is **"9"** in the url `http://swathub.com/app/support/samples/scenarios/set/9`
* `tags`: tags filtering the scenarios to export, separated by comma.
* `platforms`: the list of the platforms to export
* `status`: the status of the result to export, `finished`, `failed`, `ok` or `ng`

#### Sample target file

```
{
  "ids":[],
  "setID": "1",
  "tags":"test1",
  "status": "finished",
  "platforms": ["Mac OSX + Firefox", "Mac OSX + Chrome"]
}
```
