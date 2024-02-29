(function () {
    "use strict";

    let messageBanner;

    Office.onReady(function () {
        $(document).ready(function () {
            // Initialize the FabricUI notification mechanism and hide it.
            const element = document.querySelector('.MessageBanner');
            messageBanner = new components.MessageBanner(element);
            messageBanner.hideBanner();

            $('#insert-image').on("click", insertImage);
            $('#insert-text').on("click", insertText);
            $('#get-slide-metadata').on("click", getSlideMetadata);
            $('#add-slides').on("click", addSlides);
            $('#go-to-first-slide').on("click", goToFirstSlide);
            $('#go-to-next-slide').on("click", goToNextSlide);
            $('#go-to-previous-slide').on("click", goToPreviousSlide);
            $('#go-to-last-slide').on("click", goToLastSlide);
        });
    });

    function insertImage() {
        // Get image from web service (as a Base64-encoded string).
        $.ajax({
            url: "/api/photo/",
            dataType: "text",
            success: function (result) {
                insertImageFromBase64String(result);
            }, error: function (xhr, status, error) {
                showNotification("Error", "Oops, something went wrong.");
            }
        });
    }

    function insertImageFromBase64String(image) {
        // Call Office.js to insert the image into the document.
        Office.context.document.setSelectedDataAsync(image, {
            coercionType: Office.CoercionType.Image
        },
        function (asyncResult) {
            if (asyncResult.status === Office.AsyncResultStatus.Failed) {
                showNotification("Error", asyncResult.error.message);
            }
        });
    }

    function insertText() {
        Office.context.document.setSelectedDataAsync('Hello World!',
            function (asyncResult) {
                if (asyncResult.status === Office.AsyncResultStatus.Failed) {
                    showNotification("Error", asyncResult.error.message);
                }
            });
    }

    function getSlideMetadata() {
        Office.context.document.getSelectedDataAsync(Office.CoercionType.SlideRange,
            function (asyncResult) {
                if (asyncResult.status === Office.AsyncResultStatus.Failed) {
                    showNotification("Error", asyncResult.error.message);
                } else {
                    showNotification("Metadata for selected slide(s):", JSON.stringify(asyncResult.value), null, 2);
                }
            }
        );
    }

    async function addSlides() {
        await PowerPoint.run(async function (context) {
            context.presentation.slides.add();
            context.presentation.slides.add();

            await context.sync();

            showNotification("Success", "Slides added.");
            goToLastSlide();
        });
    }

    function goToFirstSlide() {
        Office.context.document.goToByIdAsync(Office.Index.First, Office.GoToType.Index,
            function (asyncResult) {
                if (asyncResult.status == "failed") {
                    showNotification("Error", asyncResult.error.message);
                }
            });
    }

    function goToLastSlide() {
        Office.context.document.goToByIdAsync(Office.Index.Last, Office.GoToType.Index,
            function (asyncResult) {
                if (asyncResult.status == "failed") {
                    showNotification("Error", asyncResult.error.message);
                }
            });
    }

    function goToPreviousSlide() {
        Office.context.document.goToByIdAsync(Office.Index.Previous, Office.GoToType.Index,
            function (asyncResult) {
                if (asyncResult.status == "failed") {
                    showNotification("Error", asyncResult.error.message);
                }
            });
    }

    function goToNextSlide() {
        Office.context.document.goToByIdAsync(Office.Index.Next, Office.GoToType.Index,
            function (asyncResult) {
                if (asyncResult.status == "failed") {
                    showNotification("Error", asyncResult.error.message);
                }
            });
    }

    // Helper function for displaying notifications.
    function showNotification(header, content) {
        $("#notification-header").text(header);
        $("#notification-body").text(content);
        messageBanner.showBanner();
        messageBanner.toggleExpansion();
    }
})();