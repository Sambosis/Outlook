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
                $('#results').empty();

                if (!response.results || response.results.length === 0) {
                    $('#results').html('<p class="text-center">No results found.</p>');
                    return;
                }

                    response.results.forEach(function (item) {
                        let resultHtml = `
                            <div class="result-item" data-path="${item.path}">
                                <h5 class="mb-1">${item.name}</h5>
                                <p class="email-snippet mb-0">${item.snippet}</p>
                            </div>
                        `;
                        $('#results').append(resultHtml);
                    });
                },
                error: function (error) {
                    console.error('Search error:', error); // Debug log
                    $('#results').html(`
                        <div class="alert alert-danger">
                            Error performing search: ${error}
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
                $('.result-item').removeClass('active');
                $clickedItem.addClass('active');
                // Load attachments after email content is loaded
                let emailId = filePath.replace('.html', '');
                
                displayAttachments(emailId);
            },
            error: function (xhr, status, error) {
                console.error('Error loading email:', error);
                $('#email-content').html('<div class="alert alert-danger"> Error loading email content. Please try again. </div>');
                $('#attachment-list').html('');
            }
        });
    });





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