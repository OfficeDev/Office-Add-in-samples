/* 
* Copyright (c) Microsoft. All rights reserved. Licensed under the MIT license. 
* See LICENSE in the project root for license information. 
*/ 

// Helper based on Office UI Fabric dropdown, which hides the original 'select' dropdown and 
// creates a "fake" dropdown that can be more easily styled across browsers.
// http://dev.office.com/fabric/components/dropdown
function useFabricDropdown (id) {
    const dropdownWrapper = document.getElementById(id);
    const originalDropdown = dropdownWrapper.querySelector('.ms-Dropdown-select');
    const originalDropdownOptions = originalDropdown.querySelectorAll('option');
    let newDropdownTitle = '';
    let newDropdownItems = '';
    let newDropdownSource = '';

    /** Go through the options to fill up newDropdownTitle and newDropdownItems. */
    originalDropdownOptions.forEach((option, index) => {

        /** If the option is selected, it should be the new dropdown's title. */
        if (option.selected) {
            newDropdownTitle = option.textContent;
        }

        /** Add this option to the list of items. */
        newDropdownItems += '<li class="ms-Dropdown-item' + ( (option.disabled) ? ' is-disabled"' : '"' ) + '>' + option.textContent + '</li>';
    
    });

    /** Insert the replacement dropdown. */
    newDropdownSource = '<span class="ms-Dropdown-title">' + newDropdownTitle + '</span><ul class="ms-Dropdown-items">' + newDropdownItems + '</ul>';
    dropdownWrapper.insertAdjacentHTML('beforeend', newDropdownSource);

    function _openDropdown(evt) {
        if (!dropdownWrapper.classList.contains('is-disabled')) {

            /** First, let's close any open dropdowns on this page. */
            const openDropdowns = dropdownWrapper.querySelectorAll('.is-open');
            openDropdowns.forEach(el => el.classList.remove('is-open'));

            /** Stop the click event from propagating, which would just close the dropdown immediately. */
            evt.stopPropagation();

            /** Before opening, size the items list to match the dropdown. */
            const dropdownWidth = dropdownWrapper.offsetWidth;
            const itemsList = dropdownWrapper.querySelector('.ms-Dropdown-items');
            if (itemsList) {
                itemsList.style.width = dropdownWidth + 'px';
            }
        
            /** Go ahead and open that dropdown. */
            dropdownWrapper.classList.toggle('is-open');
            
            // Close other dropdowns
            const allDropdowns = document.querySelectorAll('.ms-Dropdown');
            allDropdowns.forEach(dropdown => {
                if (dropdown !== dropdownWrapper) {
                    dropdown.classList.remove('is-open');
                }
            });

            /** Temporarily bind an event to the document that will close this dropdown when clicking anywhere. */
            function closeDropdown() {
                dropdownWrapper.classList.remove('is-open');
                document.removeEventListener('click', closeDropdown);
            }
            document.addEventListener('click', closeDropdown);
        }
    }

    /** Toggle open/closed state of the dropdown when clicking its title. */
    dropdownWrapper.addEventListener('click', function(event) {
        if (event.target.classList.contains('ms-Dropdown-title')) {
            _openDropdown(event);
        }
    });

    /** Keyboard accessibility */
    dropdownWrapper.addEventListener('keyup', function(event) {
        const keyCode = event.keyCode || event.which;
        // Open dropdown on enter or arrow up or arrow down and focus on first option
        if (!this.classList.contains('is-open')) {
            if (keyCode === 13 || keyCode === 38 || keyCode === 40) {
                _openDropdown(event);
                const firstItem = this.querySelector('.ms-Dropdown-item');
                if (firstItem && !this.querySelector('.ms-Dropdown-item.is-selected')) {
                    firstItem.classList.add('is-selected');
                }
            }
        }
        else if (this.classList.contains('is-open')) {
            const selectedItem = this.querySelector('.ms-Dropdown-item.is-selected');
            // Up arrow focuses previous option
            if (keyCode === 38 && selectedItem) {
                const prevItem = selectedItem.previousElementSibling;
                if (prevItem && prevItem.classList.contains('ms-Dropdown-item')) {
                    selectedItem.classList.remove('is-selected');
                    prevItem.classList.add('is-selected');
                }
            }
            // Down arrow focuses next option
            if (keyCode === 40 && selectedItem) {
                const nextItem = selectedItem.nextElementSibling;
                if (nextItem && nextItem.classList.contains('ms-Dropdown-item')) {
                    selectedItem.classList.remove('is-selected');
                    nextItem.classList.add('is-selected');
                }
            }
            // Enter to select item
            if (keyCode === 13) {
                if (!dropdownWrapper.classList.contains('is-disabled') && selectedItem) {

                    // Item text
                    const selectedItemText = selectedItem.textContent;

                    const titleElement = this.querySelector('.ms-Dropdown-title');
                    if (titleElement) {
                        titleElement.textContent = selectedItemText;
                    }

                    /** Update the original dropdown. */
                    const options = originalDropdown.querySelectorAll('option');
                    options.forEach(option => {
                        if (option.textContent === selectedItemText) {
                            option.selected = true;
                        } else {
                            option.selected = false;
                        }
                    });
                    
                    // Trigger change event
                    const changeEvent = new Event('change', { bubbles: true });
                    originalDropdown.dispatchEvent(changeEvent);

                    this.classList.remove('is-open');
                }
            }                
        }

        // Close dropdown on esc
        if (keyCode === 27) {
            this.classList.remove('is-open');
        }
    });

    /** Select an option from the dropdown. */
    dropdownWrapper.addEventListener('click', function(event) {
        if (event.target.classList.contains('ms-Dropdown-item')) {
            const clickedItem = event.target;
            
            if (!dropdownWrapper.classList.contains('is-disabled')) {

                /** Deselect all items and select this one. */
                const allItems = dropdownWrapper.querySelectorAll('.ms-Dropdown-item');
                allItems.forEach(item => item.classList.remove('is-selected'));
                clickedItem.classList.add('is-selected');

                /** Update the replacement dropdown's title. */
                const titleElement = dropdownWrapper.querySelector('.ms-Dropdown-title');
                if (titleElement) {
                    titleElement.textContent = clickedItem.textContent;
                }

                /** Update the original dropdown. */
                const selectedItemText = clickedItem.textContent;
                const options = originalDropdown.querySelectorAll('option');
                options.forEach(option => {
                    if (option.textContent === selectedItemText) {
                        option.selected = true;
                    } else {
                        option.selected = false;
                    }
                });
                
                // Trigger change event
                const changeEvent = new Event('change', { bubbles: true });
                originalDropdown.dispatchEvent(changeEvent);
            }
        }
    });
}