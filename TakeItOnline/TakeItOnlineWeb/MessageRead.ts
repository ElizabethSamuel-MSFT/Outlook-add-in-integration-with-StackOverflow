declare var fabric: any;

(function () {
    "use strict";

    // The Office initialize function must be run each time a new page is loaded.
    Office.initialize = function (reason) {
        $(document).ready(function () {
            $("#getMessage").click(() => tryCatch(loadMessage));

        });
    };

    // Load properties from the Item base object, then load the
    // message-specific properties.
    async function loadMessage() {
        var item = Office.context.mailbox.item;
        var readMessageItem = item as Office.MessageRead;
        $('#sender').text(readMessageItem.from.emailAddress);
        await readMessageItem.body.getAsync(Office.CoercionType.Text, function (result) {
            $('#message').text(result.value);
        });
    }

    /** Default helper for invoking an action and handling errors. */
    async function tryCatch(callback) {
        try {
            await callback();
        }
        catch (error) {
            console.log(error);
        }
    }
})();