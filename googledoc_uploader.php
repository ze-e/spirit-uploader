<?php
/*
Plugin Name: Upload Doc to Post
Description: Converts uploaded Word documents into WordPress posts with the category "story".
Version: 2.0
Author: Zachary Rex Rodriguez
*/

require 'vendor/autoload.php'; // Ensure this path points to your Composer autoload.php

use PhpOffice\PhpWord\IOFactory;

// Add a menu item to the admin menu
add_action('admin_menu', 'udtp_add_admin_menu');

function udtp_add_admin_menu() {
    add_menu_page('Upload Doc to Post', 'Upload Doc to Post', 'manage_options', 'upload-doc-to-post', 'udtp_options_page');
}

// Display the plugin options page
function udtp_options_page() {
    ?>
    <div class="wrap">
        <h2>Upload Doc to Post</h2>
        <div id="status-message">
            <?php if (isset($_GET['status']) && $_GET['status'] == 'success') : ?>
                <div style="color: green; font-weight: bold;">Post(s) created successfully!</div>
            <?php elseif (isset($_GET['status']) && $_GET['status'] == 'error') : ?>
                <div style="color: red; font-weight: bold;">One or more posts were not successful :(</div>
            <?php endif; ?>
        </div>
        <form id="udtp-form" method="post" enctype="multipart/form-data" action="<?php echo admin_url('admin-post.php'); ?>">
            <input type="hidden" name="action" value="udtp_import">
            <table class="form-table">
                <?php for ($i = 1; $i <= 10; $i++) : ?>
                    <tr valign="top">
                        <th scope="row">Upload Word Document <?php echo $i; ?></th>
                        <td><input type="file" name="udtp_doc_file_<?php echo $i; ?>" class="udtp-file-input" /></td>
                    </tr>
                <?php endfor; ?>
                <tr valign="top">
                    <th scope="row">Enable Verbose Logging</th>
                    <td><input type="checkbox" name="udtp_enable_log" value="1" /></td>
                </tr>
            </table>
            <?php submit_button('Import Documents', 'primary', 'import-button'); ?>
        </form>
        <script>
            document.addEventListener('DOMContentLoaded', function() {
                const form = document.getElementById('udtp-form');
                const importButton = document.getElementById('import-button');
                const fileInputs = document.querySelectorAll('.udtp-file-input');
                const statusMessage = document.getElementById('status-message');

                function checkFiles() {
                    let hasFile = false;
                    fileInputs.forEach(input => {
                        if (input.files.length > 0) {
                            hasFile = true;
                        }
                    });
                    importButton.disabled = !hasFile;
                }

                fileInputs.forEach(input => {
                    input.addEventListener('change', checkFiles);
                });

                form.addEventListener('submit', function(e) {
                    importButton.disabled = true;
                    statusMessage.innerHTML = '<div style="color: blue; font-weight: bold;">Loading...</div>';
                });

                checkFiles();
            });
        </script>
    </div>
    <?php
}

// Handle the form submission to import the uploaded documents
add_action('admin_post_udtp_import', 'udtp_import_uploaded_docs');

function log_message($message, $enable_log = false) {
    if ($enable_log && defined('WP_DEBUG') && WP_DEBUG) {
        error_log($message);
    }
}

function udtp_import_uploaded_docs() {
    $enable_log = isset($_POST['udtp_enable_log']) && $_POST['udtp_enable_log'] == '1';
    $successful_posts = [];

    // Iterate through each possible file input
    for ($i = 1; $i <= 10; $i++) {
        $file_key = 'udtp_doc_file_' . $i;
        if (isset($_FILES[$file_key]) && $_FILES[$file_key]['error'] == UPLOAD_ERR_OK) {
            log_message('File uploaded successfully: ' . $_FILES[$file_key]['name'], $enable_log);
            $file = $_FILES[$file_key]['tmp_name'];
            $post_title = format_title_from_filename($_FILES[$file_key]['name']);
            $content = udtp_get_doc_content($file, $enable_log);

            if ($content) {
                if (term_exists('story', 'category') === 0) {
                    wp_insert_term('story', 'category');
                }

                $post_id = wp_insert_post(array(
                    'post_title' => $post_title,
                    'post_content' => $content,
                    'post_status' => 'draft',
                    'post_category' => array(get_cat_ID('story')),
                ));

                if ($post_id) {
                    log_message('Post created successfully: ' . $post_id, $enable_log);
                    $successful_posts[] = $post_id;
                } else {
                    error_log('Failed to create post: ' . $post_title);
                }
            } else {
                error_log('Failed to get content from document: ' . $_FILES[$file_key]['name']);
            }
        } elseif (isset($_FILES[$file_key]) && $_FILES[$file_key]['error'] != UPLOAD_ERR_NO_FILE) {
            error_log('File upload error: ' . $_FILES[$file_key]['error']);
        }
    }

    if (count($successful_posts) > 1) {
        wp_redirect(admin_url('edit.php'));
    } elseif (count($successful_posts) == 1) {
        wp_redirect(get_edit_post_link($successful_posts[0], ''));
    } else {
        wp_redirect(admin_url('admin.php?page=upload-doc-to-post&status=error'));
    }
    exit;
}

// Fetch the content of the uploaded document and handle images
function udtp_get_doc_content($file, $enable_log) {
    if (class_exists('ZipArchive')) {
        log_message('ZipArchive class is available', $enable_log);
    } else {
        error_log('ZipArchive class is NOT available');
    }

    try {
        $phpWord = IOFactory::load($file, 'Word2007');
        $content = '';
        $elements = [];
        $images = [];
        $uploadDir = wp_upload_dir();

        // Iterate through sections and elements to extract content and images
        foreach ($phpWord->getSections() as $section) {
            foreach ($section->getElements() as $element) {
                if (method_exists($element, 'getText')) {
                    $elements[] = ['type' => 'text', 'content' => $element->getText()];
                }
                if (method_exists($element, 'getElements')) {
                    foreach ($element->getElements() as $innerElement) {
                        if (method_exists($innerElement, 'getText')) {
                            $elements[] = ['type' => 'text', 'content' => $innerElement->getText()];
                        }
                        if (get_class($innerElement) === 'PhpOffice\PhpWord\Element\Image') {
                            $source = $innerElement->getSource();
                            log_message('Found an image: ' . $source, $enable_log);
                            $imageData = file_get_contents($source);
                            if ($imageData === false) {
                                error_log('Failed to get image data from: ' . $source);
                                continue;
                            }
                            $filename = basename($source);
                            $filetype = wp_check_filetype($filename, null);
                            $filepath = $uploadDir['path'] . '/' . $filename;

                            if (file_put_contents($filepath, $imageData)) {
                                $attachment = array(
                                    'guid' => $uploadDir['url'] . '/' . basename($filename),
                                    'post_mime_type' => $filetype['type'],
                                    'post_title' => sanitize_file_name($filename),
                                    'post_content' => '',
                                    'post_status' => 'inherit'
                                );

                                $attach_id = wp_insert_attachment($attachment, $filepath);
                                require_once(ABSPATH . 'wp-admin/includes/image.php');
                                $attach_data = wp_generate_attachment_metadata($attach_id, $filepath);
                                wp_update_attachment_metadata($attach_id, $attach_data);

                                $imageUrl = wp_get_attachment_url($attach_id);
                                $elements[] = ['type' => 'image', 'content' => '<img src="' . $imageUrl . '" alt="' . $filename . '" />'];
                            } else {
                                error_log('Failed to write image data to: ' . $filepath);
                            }
                        }
                    }
                }
            }
        }

        // Combine text and images in the order they appear in the document
        foreach ($elements as $element) {
            $content .= $element['content'] . "\n";
        }

        return $content;
    } catch (Exception $e) {
        error_log('Exception: ' . $e->getMessage());
        return false;
    }
}

// Helper function to format the title from the file name
function format_title_from_filename($filename) {
    // List of words to lowercase
    $lowercase_words = ['a', 'an', 'and', 'as', 'at', 'but', 'by', 'for', 'in', 'nor', 'of', 'on', 'or', 'the', 'to', 'up'];

    // Remove file extension
    $title = pathinfo($filename, PATHINFO_FILENAME);

    // Replace underscores and hyphens with spaces
    $title = str_replace(['_', '-'], ' ', $title);

    // Capitalize words
    $words = explode(' ', $title);
    foreach ($words as &$word) {
        if (!in_array(strtolower($word), $lowercase_words)) {
            $word = ucfirst($word);
        }
    }

    // Always capitalize the first word
    $words[0] = ucfirst($words[0]);

    return implode(' ', $words);
}
?>
