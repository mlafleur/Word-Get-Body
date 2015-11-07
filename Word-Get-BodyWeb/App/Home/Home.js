
(function () {
    "use strict";

    // The initialize function must be run each time a new page is loaded
    Office.initialize = function (reason) {
        $(document).ready(function () {
            app.initialize();
            $('#get-body-text').click(getBodyText);
        });
    };

    function getBodyText() {
        Word.run(function (context) {

            // Get the current document's body
            var body = context.document.body;

            // Request the text from the document's body. 
            context.load(body, 'text');

            // Request the HTML representation of the body as well
            var bodyHTML = body.getHtml();

            // Synchronize the document state by executing the queued commands, 
            // and return once the task has task completion.
            return context.sync().then(function () {

                // Add the raw text to the docText <div/>
                $('#docText').html(body.text);

                // Add the HTML render to the docHtml <div/>
                $('#docHtml').html(bodyHTML.value);
            });
        })
            .catch(function (error) {
                console.log('Error: ' + JSON.stringify(error));
                if (error instanceof OfficeExtension.Error) {
                    console.log('Debug info: ' + JSON.stringify(error.debugInfo));
                }
            });
    }
})();