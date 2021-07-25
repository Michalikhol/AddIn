
var displayLenguage
var UIText

 Office.initialize = function () {

    displayLenguage = Office.context.displayLanguage;
    UIText = UIStrings.getLocaleStrings(displayLenguage);

    $("#title").text(UIText.Title)
    $("#message").text(UIText.message);
    $("#sendButton").text(UIText.sendB);
    $("#DontSendButton").text(UIText.dontSendB);

};


/* Store the locale-specific strings */

var UIStrings = (function ()
{
    "use strict";

    var UIStrings = {};

    // JSON object for English strings
    UIStrings.EN =
    {
        "Title": "Classification",
        "message": "We found recipients with different classification. Are you sure you want to send the message?",
        "sendB": "Send",
        "dontSendB": "Don't Send"
    };

    // JSON object for Hebrow strings
    UIStrings.HE =
    {
        "Title": "סיווג",
        "message": "נמצאו נמנים בסיווג שונה. האם לשלוח ?",
        "sendB": "שליחה",
        "dontSendB": "ביטול"
    };

    UIStrings.getLocaleStrings = function (locale)
    {
        var text;

        // Get the resource strings that match the language.
        switch (locale)
        {
            case 'en-US':
                text = UIStrings.EN;
                break;
            case 'he-IL':
                text = UIStrings.HE;
                break;
            default:
                text = UIStrings.EN;
                break;
        }

        return text;
    };

    return UIStrings;
})();