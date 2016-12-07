##  Result Export Tool

A desktop tool to export a single test result into different formats.
* `excel`: an Excel file with all the screenshots and parameters
* `html`: a zip file with all the page HTML files
* `diag`: a zip file with all evidences and page knowledges for diagnosis

### Build

`make.sh`

### Usage

`java -jar ResultExport.jar <path to config file> <resultFormat> <resultID>`

* `resultFormat`: type of the result format, `excel`, `html`, or `diag`
* `resultID`: the unique ID of the test result. For instance, the result ID is **12345** in the url `http://swathub.com/app/support/samples/results/12345`

**Hint**: if the size of the test result is quite large, please adjust the JVM arguments to increase the heap size. For instance:

`java -Xms1024m -Xmx1024m -jar ResultExport.jar <path to config file> <resultFormat> <resultID>`

### Config file

#### Config parameters

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
