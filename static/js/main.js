// static/js/main.js
$(document).ready(function () {
    let lastQuery = '';

    $('#search-form').on('submit', function (e) {
        e.preventDefault();
        $('#results').removeClass('d-none');
        $('.recent-emails').addClass('d-none');
        let query = $('#search-query').val().trim();

        if (query === '') {
            alert('Please enter a search query.');
            return;
        }

        if (query === lastQuery) {
            return; // Prevent duplicate searches
        }

        lastQuery = query;

        $('#results').html('<p class="text-center"><i>Searching...</i></p>');
        $('#email-content').html('<p>Select an email from the search results to view its content.</p>');

        $.ajax({
            url: '/search',
            method: 'GET',
            data: { query: query },
            success: function (response) {
                console.log('Search response:', response); // Debug log
                $('#results').empty(); // Clear previous results or "Searching..." message

                if (response.error) {
                    console.error('Search error from server:', response.error);
                    $('#results').html(`<div class="alert alert-danger">Error: ${response.error}</div>`);
                    return;
                }

                if (!response.results || response.results.length === 0) {
                    $('#results').html('<p class="text-center">No results found.</p>');
                    return;
                }

                // Update search results display using the new fields
                response.results.forEach(function (item) {
                    // Use the structure consistent with recent emails in index.html
                    let resultHtml = `
                        <div class="list-group-item result-item" data-path="${item.path}">
                            <h5 class="mb-1">${item.subject}</h5>
                            <p class="mb-1"><strong>From:</strong> ${item.sender}</p>
                            <small class="text-muted">${item.datetime_received}</small>
                        </div>
                    `;
                    $('#results').append(resultHtml);
                });
            },
            error: function (xhr, status, error) { // Use xhr, status, error for more details
                console.error('Search AJAX error:', status, error); // Debug log
                let errorMsg = `Error performing search: ${error}`;
                if (xhr.responseJSON && xhr.responseJSON.error) {
                    errorMsg = `Error performing search: ${xhr.responseJSON.error}`;
                } else if (xhr.responseText) {
                    // Try to show more specific error if available
                    errorMsg = `Error performing search: ${xhr.status} ${xhr.statusText}`;
                }
                $('#results').html(`
                    <div class="alert alert-danger">
                        ${errorMsg}
                    </div>
                `);
            }
        });
    });

    // Delegate event handler for result items (works for dynamically added elements)
    $(document).on('click', '.result-item', function () {
        let $clickedItem = $(this);
        let filePath = $(this).data('path');
        $('#email-content').html('<p class="text-center"><i>Loading...</i></p>');
        $('#attachment-list').html('<p class="text-center"><i>Loading attachments...</i></p>');
        // Load email content
        $.ajax({
            url: '/view/' + encodeURIComponent(filePath),
            method: 'GET',
            success: function (data) {
                $('#email-content').html(data);
                $('.result-item').removeClass('active'); // Ensure active class is managed correctly
                $clickedItem.addClass('active');
                // Load attachments after email content is loaded
                displayAttachments(filePath); // Pass the full relative path
            },
            error: function (xhr, status, error) {
                console.error('Error loading email:', error);
                $('#email-content').html('<p class="text-danger">Error loading email content.</p>');
            }
        });
    });

    // Function to display attachments
    function displayAttachments(emailBase) {
        $.ajax({
            url: '/list-attachments/' + encodeURIComponent(emailBase),
            method: 'GET',
            dataType: 'json',
            success: function (response) {
                if (response.attachments && response.attachments.length > 0) {
                    let attachmentsHtml = '<h4>Attachments:</h4><ul>';
                    response.attachments.forEach(function (attachment) {
                        attachmentsHtml += `<li><a href="${attachment.path}" download>${attachment.filename}</a></li>`;
                    });
                    attachmentsHtml += '</ul>';
                    $('#attachment-list').html(attachmentsHtml);
                } else {
                    $('#attachment-list').html('<p>No attachments found.</p>');
                }
            },
            error: function (xhr, status, error) {
                console.error('Error fetching attachments:', error);
                $('#attachment-list').html('<p>Error loading attachments.</p>');
            }
        });
    }

    // Add a way to go back to recent emails
    $('<button>')
        .addClass('btn btn-link btn-sm mb-2')
        .text('← Back to Recent Emails')
        .prependTo('#results')
        .click(function() {
            $('#results').addClass('d-none');
            $('.recent-emails').removeClass('d-none');
        });
});