document.addEventListener('DOMContentLoaded', function() {
    const fileInput = document.getElementById('docx-file');
    const fileNameDisplay = document.getElementById('file-name');
    const convertBtn = document.getElementById('convert-btn');
    const downloadBtn = document.getElementById('download-btn');
    const copyBtn = document.getElementById('copy-btn');
    const markdownResult = document.getElementById('markdown-result');
    
    let selectedFile = null;
    let convertedMarkdown = '';
    
    // Handle file selection
    fileInput.addEventListener('change', function(event) {
        selectedFile = event.target.files[0];
        
        if (selectedFile) {
            fileNameDisplay.textContent = selectedFile.name;
            convertBtn.disabled = false;
            markdownResult.textContent = 'Your converted markdown will appear here...';
            downloadBtn.disabled = true;
            copyBtn.disabled = true;
        } else {
            fileNameDisplay.textContent = 'No file selected';
            convertBtn.disabled = true;
        }
    });
    
    // Handle conversion
    convertBtn.addEventListener('click', function() {
        if (!selectedFile) return;
        
        // Show loading state
        convertBtn.disabled = true;
        markdownResult.textContent = 'Converting...';
        
        const reader = new FileReader();
        
        reader.onload = function(loadEvent) {
            const arrayBuffer = loadEvent.target.result;
            
            // Step 1: Use mammoth.js to convert docx to HTML
            mammoth.convertToHtml({ arrayBuffer: arrayBuffer })
                .then(function(result) {
                    const html = result.value;
                    
                    // Step 2: Use TurndownService to convert HTML to Markdown
                    const turndownService = new TurndownService({
                        headingStyle: 'atx',
                        codeBlockStyle: 'fenced',
                        emDelimiter: '*'
                    });
                    
                    // Add image handling rule
                    turndownService.addRule('images', {
                        filter: 'img',
                        replacement: function(content, node) {
                            const src = node.getAttribute('src') || '';
                            const alt = node.getAttribute('alt') || '';
                            const title = node.getAttribute('title') || '';
                            
                            // Check if it's a base64 encoded image
                            if (src.startsWith('data:image/')) {
                                // Extract image type (e.g., png, jpeg)
                                const imageType = src.split(';')[0].split('/')[1];
                                
                                // Create a placeholder instead of including the entire base64 string
                                return `![${alt}](Image: ${imageType} image${title ? ` "${title}"` : ''})`;
                            }
                            
                            // Regular image with URL
                            return `![${alt}](${src}${title ? ` "${title}"` : ''})`;
                        }
                    });
                    
                    // Improved table conversion rule
                    turndownService.addRule('tables', {
                        filter: 'table',
                        replacement: function(content, node) {
                            // Get all rows from the table
                            const rows = node.querySelectorAll('tr');
                            if (rows.length === 0) return '';
                            
                            let markdownTable = '\n\n';
                            
                            // Process each row
                            rows.forEach((row, rowIndex) => {
                                const cells = row.querySelectorAll('th, td');
                                let rowContent = '| ';
                                
                                // Process each cell in the row
                                cells.forEach(cell => {
                                    // Get cell content and clean it
                                    let cellContent = cell.textContent.trim();
                                    // Replace any pipe characters in the content
                                    cellContent = cellContent.replace(/\|/g, '\\|');
                                    rowContent += cellContent + ' | ';
                                });
                                
                                markdownTable += rowContent + '\n';
                                
                                // Add separator row after the header row
                                if (rowIndex === 0) {
                                    let separator = '| ';
                                    cells.forEach(() => {
                                        separator += '--- | ';
                                    });
                                    markdownTable += separator + '\n';
                                }
                            });
                            
                            return markdownTable + '\n';
                        }
                    });
                    
                    convertedMarkdown = turndownService.turndown(html);
                    markdownResult.textContent = convertedMarkdown;
                    
                    // Enable download and copy buttons
                    downloadBtn.disabled = false;
                    copyBtn.disabled = false;
                    convertBtn.disabled = false;
                })
                .catch(function(error) {
                    markdownResult.textContent = 'Error converting file: ' + error.message;
                    convertBtn.disabled = false;
                });
        };
        
        reader.onerror = function() {
            markdownResult.textContent = 'Error reading file';
            convertBtn.disabled = false;
        };
        
        reader.readAsArrayBuffer(selectedFile);
    });
    
    // Handle download
    downloadBtn.addEventListener('click', function() {
        if (!convertedMarkdown) return;
        
        const blob = new Blob([convertedMarkdown], { type: 'text/markdown' });
        const url = URL.createObjectURL(blob);
        const a = document.createElement('a');
        
        // Create filename from original file
        const originalName = selectedFile.name;
        const markdownName = originalName.replace(/\.docx$/, '.md');
        
        a.href = url;
        a.download = markdownName;
        document.body.appendChild(a);
        a.click();
        
        // Clean up
        setTimeout(function() {
            document.body.removeChild(a);
            window.URL.revokeObjectURL(url);
        }, 0);
    });
    
    // Handle copy to clipboard
    copyBtn.addEventListener('click', function() {
        if (!convertedMarkdown) return;
        
        navigator.clipboard.writeText(convertedMarkdown)
            .then(function() {
                const originalText = copyBtn.textContent;
                copyBtn.textContent = 'Copied!';
                
                setTimeout(function() {
                    copyBtn.textContent = originalText;
                }, 2000);
            })
            .catch(function(err) {
                console.error('Failed to copy text: ', err);
            });
    });
}); 