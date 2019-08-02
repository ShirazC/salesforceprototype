# Salesforce Prototype

This project is built to allow users to download a Word document, edit the document with all of the functionality given by Microsoft Word, and then uploading it back to Salesforce with the click of a button.

## Getting Started

This document will give you a run down of the process of getting the document from Salesforce, to you local computer, and back to Salesforce. 

### *Important*

Upon opening the Word Document, a script will run, potentially changing the document. This script's functionality is detailed down below, but this is a warning that your document may change as a result of opening it.

### Prerequisites

You will need Microsoft Word to be able to work with this project, as the macro that sends document back to Salesforce is written for Microsoft Word and already embedded in the document upon download, along with information to get it where it needs to be.

You will also need to set up a Connected App within your scratch org. The Connected App must have OAuth enabled. From there, you must give basic application permissions, relax IP restrictions, and allow users to self-authorize. In testing, we gave all permissions, but not all are needed in order to carry out full functionality. The Connected App is important, as it allows us to communicate back to Salesforce once we are done editing. From here, we need the following:
* Consumer Key
* Consumer Secret
* Sign-in Domain
Using these, as well as your username and password, you create a string which will be given to the VBA script later on. It will look something like this: `strURL = "https://ability-connect-8530-dev-ed.cs32.my.salesforce.com/services/oauth2/token?grant_type=password&client_id=3MVG9ic.6IFhNpPowgn.DHsehYzk3o72Q.EnVuvJTOB9v1Kr.r_jDP3_MJnMxqVGT3_f9VN0Ba6oYdPyBY5sb&client_secret=AD2A82DBCB361042FD73E40D92BDFDC3D0B7AF0F81B5DE3B5289C670B12CF35F&username=test-shrtmfd1zjtn@example.com&password=%25C6sAWXGi%5E"`.

## How It Works

### Salesforce Side

*Shiraz stuff here*

### VBA Script

#### On Opening 

Upon opening document, a script will run which fills the Title property. It selects the first line, uses that as the sanitized  text for the Title property, and then deletes the line from the document. 

This is works on the condition that the document Title field is set to the word "Title" (case sensitive). 

The reason for this script is because the Case ID (where the document will ultimately end up) is placed in the first line of this document when it is downloaded from Salesforce. It also means that once it is set, a document could be downloaded again and while retaining the same upload location.

To disable this script from running in testing or opening locally, change the Title field in the properties (File > Info > Properties) to be something other than "Title"

#### On Button Press

When the macro is activated by clicking the button it was bound to, the first code to be activated has a User Form pop up. It asks about what the document should be titled, as well as what form they would like to submit it in. These responses are then sanitized and stored for later use. Also note, you cannot submit and empty field and hitting cancel will stop the script and allow the user to continue making edits to the Word document.

Next, we create a path to store files in. Here, we make a temp folder. They are stored wherever `Environ("USERPROFILE")` returns. We also keep a path where file was downloaded to delete later. The document is saved here either as a PDF or Word document (depending on the User Input).

Next, we use a XMLHTTP object to make a POST request to the predetermined URL (talked about in Prerequisites section of the README). We take the response, which is Json sent as plain text, and input it to a library variable 'Parsed' using a Json Converter to parse the text. This allows us to access the text using a Key-Value relationship. 

Next, we make another POST request to the Instance URL given in the response from our first POST request. This contains our actual document. We use a function to convert the file to Base64 and send this along with the name of the document and Case ID from the Title property.

The last part is removing all of the files from the local computer pertaining to this document. This is done using 'Call Shell()', calling the command prompt, and running a few commands all in the same line. This is necessary because you cannot delete a file that is being used. We are able to close Microsoft Word and delete the file.

## Built With

* Visual Basic for Applications macro
* JavaScript
* Lightning Web Component

## Authors

* **Shiraz Chokshi** - *Initial work*
* **John Passalis** - *Initial work* 

## License

There is not one... yet.

## Acknowledgments
