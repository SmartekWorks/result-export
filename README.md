##  Result Export Tool

A desktop tool to export test results into different formats. Each test result will be exported as a file in one of the following formats:
* `raw`: a zip file with all the evidences
* `excel`: an Excel file with all the screenshots and parameters
* `html`: a HTML file with all the screenshots and parameters in a zip package
* `source`: a zip file with all the page HTML files
* `diag`: a zip file with all evidences and page knowledges for diagnosis

### Build

`make.sh`

### Usage

`java -jar ResultExport.jar <path to config file> <resultFormat> <path to target file>`

* `resultFormat`: type of the result format, `raw`, `excel`, `html`, `source`, or `diag`

**Hint**: if the size of the test result is quite large, please adjust the JVM arguments to increase the heap size. For instance:

`java -Xms1024m -Xmx1024m -jar ResultExport.jar <path to config file> <resultFormat> <path to target file>`

### Config file

The SWATHub API credentials are setup in this file with the following keys:

#### Parameters

* `serverUrl`: the URL of SWATHub Server URL, such as http://www.swathub.com/
* `username`: the username of SWATHub Server
* `apiKey`: the api key for the user, same as the key for robot
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

The criterias to select the target test results are defined in this file. We support two kinds of selection:
* `By ids`: the unique ID(s) of test results are defined 
* `By filter`: test set level options are provided to fetch required test restuls

#### Parameters

* `ids`: the list of the unique ID (string) of the test result. For instance, the result ID is **"12345"** in the url `http://swathub.com/app/support/samples/results/12345`.

**Note** : the following parameters will be ignored if the `ids` list is not empty.

* `lastCount`(mandatory): the result index (positive integer) in all the results meeting the filters below for any single test case. For instance, `1` means the latest result, and `2` means the one before the latest.
* `setID`(mandatory): the unique ID (string) of the test set in the target workspace, which can be got from the test set url. For instance, the set ID is **"9"** in the url `http://swathub.com/app/support/samples/scenarios/set/9`. 
* `tags`(optional): tags filtering the scenarios to export, separated by comma. 
* `platform`(optional): the platform to export. It means any platform if the value is an empty string. 
* `status`(optional): the status of the result to export, `finished`, `failed`, `ok` or `ng`. It means any status if the value is an empty string.
* `beforeDate`(optional): the date when results generated before, in the format of `YYYY/MM/DD hh:mm:ss`. It means now if the value is an empty string. Please be noted the timezone of `beforeDate` is Asia/Tokyo.

#### Sample target file

```
{
  "ids":[],
  "filters": {
    "setID": "1",
    "tags": "tag1, tag2",
    "status": "finished",
    "platform": "Windows 10 + Firefox 68",
    "beforeDate": "2017/08/08 14:00:00",
    "lastCount": 1
  }
}
```

### Supported platforms

Unlike SWATHub `v1.5.x` which defines all the supported platforms. In version `v1.6.0` and later, the platforms can be any combinations on your laptop, with the specific version of operation system and browsers. It is in the following format: 

```
<OS> <OS Version> + <Browser> <Browser Version>
```

Please check out the platform of certain task just on the result page, under the `Statistics -> Platform type`. Here're some sample platforms:
* Windows 10 + IE 11
* Windows 10 + Firefox 68
* Windows 10 + Edge 81
* mac OS 19.3.0 + Safari 13
* iOS 13.3 + Safari 13
* Android 9 + Chrome 79

The supported operation systems consist of:
* Windows, XP and above
* mac OS, 10.6 and above
* iOS, 9 and above
* Android, 5 and above

The supported browsers consist of:
* IE
* Firefox
* Chrome
* Edge
* Edge Legacy
* Safari