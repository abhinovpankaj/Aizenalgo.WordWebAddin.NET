// The initialize function must be run each time a new page is loaded.
(function () {
    Office.initialize = function (reason) {
        // If you need to initialize something you can do so here.
    };
})();

const AUTHENTICATIONBASEURL = "https://demo.aizenalgo.com:9016/api/WordProc/WordProcAuthentication";
const VERIFICATIONBASEURL = "https://demo.aizenalgo.com:9016/api/WordProc/WordProcSessionDetails";


let docProp = {};
Office.onReady(() => {
    // If needed, Office.js is ready to be called
    //readCustomDocumentProperties();

});

/**
 * Shows a notification when the add-in command is executed.
 * @param event {Office.AddinCommands.Event}
 */
function submitDocument(event) {
    try {
        console.log("Inside submitDocument function");
        readCustomDocumentProperties();
        //console.log(Office.context.document.url);


    } catch (error) {
        console.log(error);
    }

    // const message = {
    //   type: Office.MailboxEnums.ItemNotificationMessageType.InformationalMessage,
    //   message: "Submitting Docuzen document.",
    //   icon: "Icon.80x80",
    //   persistent: false,
    // };

    // Show a notification message
    //Office.context.mailbox.item.notificationMessages.replaceAsync("submitDocument", message);

    // Be sure to indicate when the add-in command function is complete
    event.completed();
}
function saveDocument(event) {
    console.log("Inside save function");

    event.completed();
}
function readCustomDocumentProperties() {
    console.log("Inside readcustom function,Commands.js");

    Word.run(async (context) => {
        //var isDocuzenDoc=false;
        const properties = context.document.properties.customProperties;
        properties.load("key,value");

        await context.sync();
        try {

            for (let i = 0; i < properties.items.length; i++) {
                if (properties.items[i].key == "DVId") {
                    //isDocuzenDoc =true;
                    docProp.dvid = properties.items[i].value;
                }
                if (properties.items[i].key == "SToken") {
                    docProp.stoken = properties.items[i].value;
                }
                if (properties.items[i].key == "Uid") {
                    docProp.uid = properties.items[i].value;
                }
                if (properties.items[i].key == "logou") {
                    docProp.logou = properties.items[i].value;
                }
            }
            //set document name and path
            var uploadFilePath = Office.context.document.url;
            var pieces = uploadFilePath.split('\\');
            var filename = pieces[pieces.length - 1];
            docProp.fileName = filename;
            docProp.uploadFile = uploadFilePath;
            console.log(docProp);

            SubmitDocumentService(docProp, 1);

        }
        catch (error) {
            console.log("read doc property:" + error.stack);
        }

    });
}
function getGlobal() {
    return typeof self !== "undefined"
        ? self
        : typeof window !== "undefined"
            ? window
            : typeof global !== "undefined"
                ? global
                : undefined;
}

const g = getGlobal();
g.submitDocument = submitDocument;
//services
function SubmitDocumentService({ stoken, dvid, uploadFile, fileName }, type) {

    const endpoint = `${VERIFICATIONBASEURL}?SessionId=${stoken}&DocID=${dvid}&Mode=${type}`;
    var dataArray = new FormData();
    //dataArray.append("fileName", fileName);
    dataArray.append("file", uploadFile);


    fetch(endpoint, {
        method: 'POST',
        body: dataArray,
        mode: 'no-cors',
        headers: {
            'access-control-allow-origin': '*'
        }
    })
        .then(response => {
            if (!response.ok) throw (`invalid response: ${response.status}`);
            return response.json()
        })
        .then(data => console.log(data))

        .catch((err) => {
            console.log(err);
        });

    //  var result= new Promise(function (resolve, reject) {
    //    fetch(endpoint,{
    //           method: 'POST',
    //           body: dataArray,       
    //           headers: {
    //             'Access-Control-Allow-Origin': '*',
    //             'Accept':'/',          
    //           }
    //        })
    //      .then(function (response){
    //        return response.json();
    //        }
    //      )
    //      .then(function (json) {
    //        resolve(JSON.stringify(json.names));
    //      })
    //  });
}