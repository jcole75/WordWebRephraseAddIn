
(function () {
    "use strict";
    var debug = false;
    var messageBanner;
    var openAIAPIKey = "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";

    //Curie model trained on single dataset with <70% word overlap. ~60,000 pairs
    var openAIModelID = "curie:ft-user-lf76r2hn1wchy9vrsd87nxj0-2021-10-06-22-17-51";

    //Babbage model trained on a sample of 3 datasets including the above (44%) of the above, QQP, and MRPC. Filtered for matching paraphrases, and filtered out profanity and suicide references
    //var openAIModelID = "babbage:ft-user-lf76r2hn1wchy9vrsd87nxj0-2021-10-16-21-03-40";

    //Curie - first finetuned simplification model
    //var openAIModelID = "curie:ft-user-lf76r2hn1wchy9vrsd87nxj0-2021-10-18-03-09-01";

    // The initialize function must be run each time a new page is loaded.
    Office.initialize = function (reason) {
        $(document).ready(function () {
            // Initialize the notification mechanism and hide it
            var element = document.querySelector('.MessageBanner');
            messageBanner = new components.MessageBanner(element);
            messageBanner.hideBanner();
            applyOfficeTheme();
            // If not using Word 2016, use fallback logic.
            if (!Office.context.requirements.isSetSupported('WordApi', '1.1')) {
                $("#template-description").text("This sample displays the selected text.");
                $('#button-text').text("Display!");
                $('#button-desc').text("Display the selected text");
                
                $('#highlight-button').click(displaySelectedText);
                return;
            }

            displaySelectedText();
            $('#button-text').text("More Choices");
            $('#button-desc').text("Use advanced AI to rephrase your writing.");

            addEventHandlers();

        });
    };

    function addEventHandlers() {
        $('#highlight-button').click(displaySelectedText);
        $('#rephrase-one-button').click(function () { replaceText(1) });
        $('#rephrase-two-button').click(function () { replaceText(2) });
    }

    function test() {
        console.log("test function");
    }

    /* write a function to replace the selected text in a microsoft word document */
    function replaceSelection(text) {
        console.log("replaceSelection(text=" + text);
        var sel = document.getSelection();
        console.log("sel=" + sel);
        var range = sel.getRangeAt(0);
        console.log("range=" + range);
        range.deleteContents();
        var textNode = document.createTextNode(text);
        range.insertNode(textNode);
    }
    
    function replaceText(btnNum) {
        console.log("replaceText btnNum=" + btnNum);
        Word.run(function (context) {
            var doc = context.document;
            var originalRange = doc.getSelection();
            var replacementText;
            if (btnNum == 1) {
                replacementText = $('#button-rephrase1-text').text().trim();
            } else if (btnNum == 2) {
                replacementText = $('#button-rephrase2-text').text().trim();
            }
            originalRange.insertText(replacementText, "Replace");
            return context.sync();
        })
            .catch(function (error) {
                console.log("Error: " + error);
                if (error instanceof OfficeExtension.Error) {
                    console.log("Debug info: " + JSON.stringify(error.debugInfo));
                }
            });
    }


    function applyOfficeTheme() {
        // Get office theme colors.
        var bodyBackgroundColor = Office.context.officeTheme.bodyBackgroundColor;
        var bodyForegroundColor = Office.context.officeTheme.bodyForegroundColor;
        var controlBackgroundColor = Office.context.officeTheme.controlBackgroundColor;
        var controlForegroundColor = Office.context.officeTheme.controlForegroundColor;
        console.log("bodyBackgroundColor=" + bodyBackgroundColor);
        // Apply body background color to a CSS class.
        //apply the office css to the pane
        $('.body').css('background-color', bodyBackgroundColor);
        $('#content-main').css('background-color', bodyBackgroundColor);
        $('.footer').css('background-color', bodyBackgroundColor);
        $('.ms-fontColor-neutralSecondary').css('color', bodyForegroundColor);
        document.body.style.backgroundColor = bodyBackgroundColor;
    }

    function displaySelectedText() {
        Office.context.document.getSelectedDataAsync(Office.CoercionType.Text,
            function (result) {
                if (result.status === Office.AsyncResultStatus.Succeeded) {
                    console.log("got selected text: " + result.value);
                    //showNotification('The selected text is:', '"' + result.value + '"');
                    if (result.value.trim().length > 0) {
                        $("#select_text_title").text("Choose the option below to replace your text or click More Choices to get another set.");
                        //showNotification("Choose the option below to replace your text or click More Choices to get another set.")
                        OpenaiFetchAPIResponse(result.value);
                        $("#template-description").text("Text to rephrase: "+result.value);
                    } else {
                        $("#select_text_title").text("No text selected.  Select text and click More Choices.");
                        //showNotification("No text selected.  Select text and click More Choices.")
                    }
                } else {
                    console.log("couldn't get selected text. Error: " + result.error.message);
                    showNotification('Error:', result.error.message);
                }
            });
    }

    function OpenaiFetchAPIResponse(sentence) {
        console.log("Calling GPT3 for sentence: " + sentence);
        console.log("openAIAPIKey " + openAIAPIKey);

        if (debug) {
            console.log("debug mode.  Not calling API");
            $('#button-rephrase1-text').text("debug text 1");
            $('#button-rephrase2-text').text("debug text 2");
        } else {
            console.log("Calling API");
            var url = "https://api.openai.com/v1/completions";
            var bearer = 'Bearer ' + openAIAPIKey
            var maxTokens = Math.round(((sentence + "\n ->").length / 4)) + 15;
            console.log("maxTokens " + maxTokens);
            //limit max tokens to 30 per OpenAI guidelines
            //https://beta.openai.com/docs/use-case-guidelines/use-case-requirements-library
            maxTokens = ((maxTokens > 30) ? 30 : maxTokens);
            console.log("maxTokens " + maxTokens);
            fetch(url, {
                method: 'POST',
                headers: {
                    'Authorization': bearer,
                    'Content-Type': 'application/json'
                },
                body: JSON.stringify({
                    "model": openAIModelID,
                    "prompt": sentence + "\n ->",

                    "max_tokens": maxTokens,
                    "temperature": 0.8,
                    "n": 2,
                    "stream": false,
                    "logprobs": null,
                    "stop": "END"
                })

            }).then(response => {

                return response.json()

            }).then(data => {
                console.log(data)
                console.log(typeof data)
                console.log(Object.keys(data))
                console.log(data['choices'][0].text)
                $('#button-rephrase1-text').text(data['choices'][0].text);
                $('#button-rephrase2-text').text(data['choices'][1].text);
            })
                .catch(error => {
                    console.log('Something bad happened ' + error)
                });
        }
    }

  
    //$$(Helper function for treating errors, $loc_script_taskpane_home_js_comment34$)$$
    function errorHandler(error) {
        // $$(Always be sure to catch any accumulated errors that bubble up from the Word.run execution., $loc_script_taskpane_home_js_comment35$)$$
        showNotification("Error:", error);
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
    }

    // Helper function for displaying notifications
    function showNotification(header, content) {
        $("#notification-header").text(header);
        $("#notification-body").text(content);
        messageBanner.showBanner();
        messageBanner.toggleExpansion();
    }
})();
