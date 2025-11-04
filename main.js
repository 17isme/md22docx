document.addEventListener('DOMContentLoaded', () => {
    const templateUploader = document.getElementById('template-uploader');
    const markdownInput = document.getElementById('markdown-input');
    const convertBtn = document.getElementById('convert-btn');
    const langSwitcher = document.getElementById('lang-switcher-select');
    
    let templateFile = null;
    let extractedStyles = null; // Stays as the style dictionary
    let analysisResult = null; // New variable to hold the full analysis
    let debounceTimer;

    // --- I18n Setup ---
    langSwitcher.addEventListener('change', (event) => setLanguage(event.target.value));
    
    // Set initial language and update switcher to match
    langSwitcher.value = currentLanguage;
    setLanguage(currentLanguage);

    templateUploader.addEventListener('change', async (event) => {
        templateFile = event.target.files[0];
        extractedStyles = null;
        analysisResult = null;
        styleListContainer.innerHTML = '<p class="placeholder">正在解析模板...</p>';
        templateXmlPreview.querySelector('code').textContent = '正在读取模板文件...';

        if (templateFile) {
            try {
                const { analysis, rawXml } = await parseTemplate(templateFile);
                analysisResult = analysis; // Store the full analysis

                // 显示原始XML
                templateXmlPreview.querySelector('code').innerHTML = highlightXml(rawXml);

                // 显示解析结果
                displayLineAnalysis(analysis);
                extractedStyles = buildStylesFromAnalysis(analysis);
            } catch (error) {
                console.error('解析模板时出错:', error);
                const errorMsg = `模板解析失败: ${error.message}`;
                styleListContainer.innerHTML = `<p class="placeholder" style="color: red;">${errorMsg}</p>`;
                templateXmlPreview.querySelector('code').textContent = errorMsg;
                alert(errorMsg);
            }
        } else {
            styleListContainer.innerHTML = '<p class="placeholder">上传模板后，此处将显示提取到的样式。</p>';
            templateXmlPreview.querySelector('code').textContent = '上传模板后，此处将显示其内部 document.xml 的原始结构...';
        }
    });

    markdownInput.addEventListener('input', () => {
        clearTimeout(debounceTimer);
        debounceTimer = setTimeout(() => {
            if (extractedStyles) {
                const markdownText = markdownInput.value;
                // No need to show preview anymore
                // const xmlContent = generateDocumentXml(markdownText, extractedStyles);
            }
        }, 300);
    });

    convertBtn.addEventListener('click', async () => {
        if (!templateFile) {
            alert(t('uploadFirst'));
            return;
        }
        if (!extractedStyles || !analysisResult) {
            alert(t('stylesNotReady'));
            return;
        }

        const markdownText = markdownInput.value;
        if (!markdownText.trim()) {
            alert(t('enterMarkdown'));
            return;
        }

        try {
            convertBtn.textContent = t('converting');
            convertBtn.disabled = true;

            const newBodyContent = generateDocumentXml(markdownText, extractedStyles);
            // Pass the full analysis to the blob creator
            const blob = await createDocxBlob(newBodyContent, templateFile, analysisResult);
            
            saveAs(blob, 'output.docx');
            alert(t('conversionSuccess'));

        } catch (error) {
            console.error('转换过程中发生错误:', error);
            alert(`${t('conversionFail')}${error.message}`);
        } finally {
            // Restore button text based on current language
            convertBtn.textContent = t('convertButton');
            convertBtn.disabled = false;
        }
    });

    function displayLineAnalysis(analysis) {
        if (!analysis || analysis.length === 0) {
            styleListContainer.innerHTML = '<p class="placeholder">在文档中没有找到任何文本内容。</p>';
            return;
        }

        styleListContainer.innerHTML = '';
        const escapeHtml = (unsafe) => unsafe.replace(/[<>&'"]/g, c => ({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;','\'':'&#039;'}[c]));

        for (const line of analysis) {
            const item = document.createElement('div');
            item.className = 'style-item';
            
            let matchLabel = '';
            if (line.match) {
                matchLabel = `<span class="match-label">[匹配为: ${line.match}]</span>`;
            }

            item.innerHTML = `<p><code>${escapeHtml(line.text)}</code> ${matchLabel}</p>`;
            styleListContainer.appendChild(item);
        }
    }

    function buildStylesFromAnalysis(analysis) {
        const styles = {};
        for (const line of analysis) {
            if (line.match && line.styleInfo && !styles[line.match]) {
                styles[line.match] = line.styleInfo;
            }
        }
        const tableAnalysis = analysis.find(item => item.match === 'table');
        if (tableAnalysis) {
            styles.table = tableAnalysis.styleInfo;
        }
        if (!styles.p) {
            styles.p = { id: 'Normal', pPr: '', rPr: '', foundText: '（默认后备）' };
        }
        return styles;
    }

    function highlightXml(xml) {
        if (typeof xml !== 'string') return '';
        
        let escapedXml = escapeXml(xml);

        // 匹配完整的标签，包括属性
        escapedXml = escapedXml.replace(/(&lt;\/?[\w:\-]+)((?:\s+[\w:\-]+=(&quot;.*?&quot;|'.*?'))*)(\s*\/?&gt;|\?&gt;)/g, (match, tag, attrs, a, endTag) => {
            let highlightedAttrs = attrs;
            // 高亮属性名和等号
            highlightedAttrs = highlightedAttrs.replace(/(\s[\w:\-]+)=/g, '<span style="color: #9cdcfe;">$1</span>=');
            // 高亮属性值
            highlightedAttrs = highlightedAttrs.replace(/(=(&quot;.*?&quot;|'.*?'))/g, '=<span style="color: #ce9178;">$1</span>');
            
            return `<span style="color: #569cd6;">${tag}</span>${highlightedAttrs}<span style="color: #569cd6;">${endTag}</span>`;
        });

        return escapedXml;
    }

    const escapeXml = (unsafe) => {
        if (typeof unsafe !== 'string') return '';
        return unsafe.replace(/[<>&'"]/g, c => ({'<':'&lt;','>':'&gt;','&':'&amp;','\'':'&apos;','"':'&quot;'}[c]));
    };

    function matchMarkdownStyle(text) {
        const styleMappings = [
            { key: 'h6', regex: /^#{6}\s*(.+)/ },
            { key: 'h5', regex: /^#{5}\s*(.+)/ },
            { key: 'h4', regex: /^#{4}\s*(.+)/ },
            { key: 'h3', regex: /^#{3}\s*(.+)/ },
            { key: 'h2', regex: /^#{2}\s*(.+)/ },
            { key: 'h1', regex: /^#{1}\s*(.+)/ },
            { key: 'blockquote', regex: /^>\s*(.+)/ },
            { key: 'ol', regex: /^\d+\.\s*(.+)/ },
            { key: 'ul', regex: /^[*\-]\s*(.+)/ },
        ];
        for (const mapping of styleMappings) {
            if (mapping.regex.test(text)) {
                return mapping.key;
            }
        }
        return null;
    }

    async function parseTemplate(file) {
        const zip = await JSZip.loadAsync(file);
        const docXmlFile = zip.file('word/document.xml');
        if (!docXmlFile) {
            throw new Error('无效的.docx文件：找不到 word/document.xml');
        }

        const rawXml = await docXmlFile.async('string');
        const xmlDoc = new DOMParser().parseFromString(rawXml, "application/xml");
        
        const analysisResult = [];
        
        // --- 1. Extract paragraph styles ---
        const paragraphs = Array.from(xmlDoc.getElementsByTagName('w:p'));
        let isDefaultParagraphTaken = false;

        for (const p of paragraphs) {
            let fullTextContent = '';
            // 遍历段落的所有直接子元素，以正确处理图文混排和分割的文本
            for (const child of Array.from(p.childNodes)) {
                if (child.nodeName === 'w:r') { // 如果是文字块
                    const textNodes = child.getElementsByTagName('w:t');
                    if (textNodes.length > 0) {
                        fullTextContent += Array.from(textNodes).map(t => t.textContent).join('');
                    }
                }
                // 可以在此添加对 <w:drawing> 等其他元素的处理
            }

            const trimmedText = fullTextContent.trim();
            if (trimmedText.length === 0) continue;

            let lineInfo = { text: trimmedText, match: null, styleInfo: null };
            const headingMatch = matchMarkdownStyle(trimmedText);

            if (headingMatch) {
                lineInfo.match = headingMatch;
            } else if (!isDefaultParagraphTaken && trimmedText.length > 0) { // Ensure it's not an empty p
                lineInfo.match = 'p';
                isDefaultParagraphTaken = true;
            }

            if (lineInfo.match) {
                const pPr = p.getElementsByTagName('w:pPr')[0];
                const styleElement = pPr?.getElementsByTagName('w:pStyle')[0];
                lineInfo.styleInfo = {
                    id: styleElement ? styleElement.getAttribute('w:val') : 'Normal',
                    pPr: pPr ? pPr.outerHTML : '',
                    rPr: p.getElementsByTagName('w:rPr')[0]?.outerHTML || '',
                    foundText: trimmedText,
                    outerHTML: p.outerHTML
                };
            }
            analysisResult.push(lineInfo);
        }

        // --- 2. Extract table styles ---
        const tableElements = Array.from(xmlDoc.getElementsByTagName('w:tbl'));
        if (tableElements.length > 0) {
            const firstTable = tableElements[0];
            const tblPr = firstTable.getElementsByTagName('w:tblPr')[0];
            const firstCell = firstTable.getElementsByTagName('w:tc')[0];
            let tcPr = null, pPrInCell = null, rPrInCell = null;
            if (firstCell) {
                tcPr = firstCell.getElementsByTagName('w:tcPr')[0];
                const pInCell = firstCell.getElementsByTagName('w:p')[0];
                if (pInCell) {
                    pPrInCell = pInCell.getElementsByTagName('w:pPr')[0];
                    rPrInCell = pInCell.getElementsByTagName('w:rPr')[0];
                }
            }
            analysisResult.push({
                text: '[模板中的第一个表格]',
                match: 'table',
                styleInfo: {
                    tblPr: tblPr ? tblPr.outerHTML : '',
                    tcPr: tcPr ? tcPr.outerHTML : '',
                    pPr: pPrInCell ? pPrInCell.outerHTML : '',
                    rPr: rPrInCell ? rPrInCell.outerHTML : '',
                    foundText: '[模板中的第一个表格]',
                    outerHTML: firstTable.outerHTML
                }
            });
        }
        
        return { analysis: analysisResult, rawXml };
    }

    function generateDocumentXml(markdown, styles) {
        const tokens = marked.lexer(markdown);
        let xmlContent = '';
        
        for (const token of tokens) {
            let style = styles.p || { id: 'Normal' };
            switch (token.type) {
                case 'heading':
                    style = styles[`h${token.depth}`] || styles.p;
                    // 只引用段落样式ID，文字样式由styles.xml定义
                    xmlContent += `<w:p><w:pPr><w:pStyle w:val="${style.id}"/></w:pPr><w:r><w:t xml:space="preserve">${escapeXml(token.text)}</w:t></w:r></w:p>`;
                    break;
                case 'paragraph':
                    style = styles.p;
                    xmlContent += `<w:p><w:pPr><w:pStyle w:val="${style.id}"/></w:pPr><w:r><w:t xml:space="preserve">${escapeXml(token.text)}</w:t></w:r></w:p>`;
                    break;
                case 'space':
                    style = styles.p;
                    xmlContent += `<w:p><w:pPr><w:pStyle w:val="${style.id}"/></w:pPr></w:p>`;
                    break;
                case 'list':
                    style = styles[token.ordered ? 'ol' : 'ul'] || styles.p;
                    for(const item of token.items) {
                        // 列表也只引用段落样式ID。具体的项目符号、缩进等应在styles.xml中定义
                        xmlContent += `<w:p><w:pPr><w:pStyle w:val="${style.id}"/></w:pPr><w:r><w:t xml:space="preserve">${escapeXml(item.text)}</w:t></w:r></w:p>`;
                    }
                    break;
                case 'blockquote':
                    style = styles.blockquote || styles.p;
                    for (const innerToken of token.tokens) {
                        if(innerToken.type === 'paragraph') {
                             xmlContent += `<w:p><w:pPr><w:pStyle w:val="${style.id}"/></w:pPr><w:r><w:t xml:space="preserve">${escapeXml(innerToken.text)}</w:t></w:r></w:p>`;
                        }
                    }
                    break;
                case 'table':
                    xmlContent += generateTableXml(token, styles);
                    break;
            }
        }
        return xmlContent;
    }

    function generateTableXml(token, styles) {
        const tableStyle = styles.table || { tblPr: '', tcPr: '', pPr: '', rPr: '' };
        const generateCellContent = (text) => `<w:p>${tableStyle.pPr}<w:r>${tableStyle.rPr}<w:t xml:space="preserve">${escapeXml(text)}</w:t></w:r></w:p>`;

        let headerRowXml = '<w:tr>';
        for (const cellText of token.header) {
            headerRowXml += `<w:tc>${tableStyle.tcPr}${generateCellContent(cellText)}</w:tc>`;
        }
        headerRowXml += '</w:tr>';

        let bodyRowsXml = '';
        for (const row of token.rows) {
            let rowXml = '<w:tr>';
            for (const cellText of row) {
                rowXml += `<w:tc>${tableStyle.tcPr}${generateCellContent(cellText)}</w:tc>`;
            }
            rowXml += '</w:tr>';
            bodyRowsXml += rowXml;
        }
        return `<w:tbl>${tableStyle.tblPr}${headerRowXml}${bodyRowsXml}</w:tbl>`;
    }

    async function createDocxBlob(newBodyContent, templateFile, analysis) {
        console.log("Surgical Strategy: Rebuilding <w:sectPr> from template references.");
        const zip = await JSZip.loadAsync(templateFile);
        let docXmlStr = await zip.file('word/document.xml').async('string');

        // Step 1: Extract critical page setup tags from the original document body.
        const bodyTagStart = docXmlStr.indexOf('<w:body');
        const bodyTagEnd = docXmlStr.indexOf('>', bodyTagStart) + 1;
        const bodyCloseTag = docXmlStr.lastIndexOf('</w:body>');
        const oldBodyContent = docXmlStr.substring(bodyTagEnd, bodyCloseTag);

        // 用正则表达式精确查找页眉/页脚引用、页面大小和边距
        const headerRefMatch = oldBodyContent.match(/<w:headerReference[^>]+r:id="[^"]+"[^>]*\/>/g) || [];
        const footerRefMatch = oldBodyContent.match(/<w:footerReference[^>]+r:id="[^"]+"[^>]*\/>/g) || [];
        const pgSzMatch = oldBodyContent.match(/<w:pgSz[^>]*\/>/);
        const pgMarMatch = oldBodyContent.match(/<w:pgMar[^>]*\/>/);
        
        const headerRefXml = headerRefMatch.join('');
        const footerRefXml = footerRefMatch.join('');
        const pgSzXml = pgSzMatch ? pgSzMatch[0] : '';
        const pgMarXml = pgMarMatch ? pgMarMatch[0] : '';

        // Step 2: 用提取到的标签重新构建 <w:sectPr>
        let sectPrInnerXml = headerRefXml + footerRefXml + pgSzXml + pgMarXml;
        let sectPrXml = '';

        if (sectPrInnerXml) {
            sectPrXml = `<w:sectPr>${sectPrInnerXml}</w:sectPr>`;
            console.log("Rebuilt sectPr:", sectPrXml);
        } else {
            console.warn("Could not find any page setup info (<w:headerReference>, <w:pgSz>, etc.) in the template.");
        }
        
        // Step 3: 组合新的body内容 = MD生成的内容 + 重建的页面设置
        const finalBodyContent = newBodyContent + sectPrXml;

        // Step 4: 用新的body内容替换掉旧的
        const finalDocXmlStr = docXmlStr.substring(0, bodyTagEnd) + finalBodyContent + docXmlStr.substring(bodyCloseTag);
        
        zip.file("word/document.xml", finalDocXmlStr);

        return zip.generateAsync({
            type: 'blob',
            mimeType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
        });
    }
});