(function () {
    "use strict";

    // The onReady function must be run each time a new page is loaded.
    Office.onReady(function (info) {
        document.getElementById("get-data-from-selection").addEventListener("click", () => getDataFromSelection());
    });

    // Gets and displays some details about the current slide.
    async function getDataFromSelection() {
        try {
            await PowerPoint.run(async (context) => {
                const slides = context.presentation.getSelectedSlides();
                slides.load("items/id,items/index");
                await context.sync();

                const details = slides.items.map((slide) => ({
                    id: slide.id,
                    index: slide.index
                }));
                document.getElementById("selected-data").textContent =
                    'Hello, world! Some slide details are: ' + JSON.stringify(details);
            });
        } catch (error) {
            document.getElementById("selected-data").textContent = 'Error getting slide details.';
            console.error('Error:', error.message);
        }
    }
})();