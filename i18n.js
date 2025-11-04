const translations = {
    en: {
        appTitle: "Markdown to DOCX Converter",
        step1Title: "Step 1: Upload Style Template (.docx)",
        step1Desc1: "In your .docx template file, please do the following:",
        step1List1: "<b>Style various texts:</b> Enter Markdown-style marker text like <code># Heading 1</code>, <code>## Heading 2</code>, <code>> Quote</code>, <code>Body Text</code>, etc., and set their respective styles.",
        step1List2: "<b>Style tables:</b> Insert a table with <b>any content</b> into the document and set your desired borders, shading, and internal text styles.",
        step1Desc2: "The program will automatically recognize these elements and use their styles as the conversion standard.",
        step2Title: "Step 2: Input Markdown Content",
        markdownPlaceholder: "Enter your Markdown here...",
        convertButton: "Convert and Download DOCX",
        footerText: `Powered by <a href="https://github.com/Stuk/jszip" target="_blank">JSZip</a>, <a href="https://github.com/markedjs/marked" target="_blank">marked.js</a>, and <a href="https://github.com/eligrey/FileSaver.js" target="_blank">FileSaver.js</a>.`,
        
        // Dynamic text from JS
        uploadFirst: "Please upload a .docx template file first!",
        stylesNotReady: "Styles have not been parsed or failed to parse. Please re-upload the template.",
        enterMarkdown: "Please enter valid Markdown content!",
        converting: "Converting...",
        conversionSuccess: "Conversion successful!",
        conversionFail: "Conversion failed: ", // Appended with error message
    },
    zh: {
        appTitle: "Markdown 转 DOCX 转换器",
        step1Title: "第一步：上传样式模板 (.docx)",
        step1Desc1: "请在您的.docx模板文件中进行如下操作：",
        step1List1: "<b>为各种文本设置样式：</b>输入<code># 标题1</code>, <code>## 标题2</code>, <code>> 引用</code>, <code>正文</code>等MD标记文本，并分别为它们设置样式。",
        step1List2: "<b>设置表格样式：</b>在文档中插入一个<b>任意内容</b>的表格，并设置好您期望的边框、底纹和内部文字样式。",
        step1Desc2: "程序会自动识别这些元素，并将它们的样式作为转换标准。",
        step2Title: "第二步：输入 Markdown 内容",
        markdownPlaceholder: "在此处输入您的 Markdown...",
        convertButton: "转换并下载 DOCX",
        footerText: `由 <a href="https://github.com/Stuk/jszip" target="_blank">JSZip</a>、<a href="https://github.com/markedjs/marked" target="_blank">marked.js</a> 和 <a href="https://github.com/eligrey/FileSaver.js" target="_blank">FileSaver.js</a> 强力驱动。`,

        // Dynamic text from JS
        uploadFirst: "请先上传一个.docx模板文件！",
        stylesNotReady: "样式尚未解析完成或解析失败，请重新上传模板。",
        enterMarkdown: "请输入有效的Markdown内容！",
        converting: "正在转换...",
        conversionSuccess: "转换成功！",
        conversionFail: "转换失败: ",
    }
};

let currentLanguage = 'en'; // Default language

function setLanguage(lang) {
    if (!translations[lang]) {
        console.warn(`Language "${lang}" not found. Defaulting to 'en'.`);
        lang = 'en';
    }
    currentLanguage = lang;
    document.documentElement.lang = lang; // Update the lang attribute of the html tag

    document.querySelectorAll('[data-i18n-key]').forEach(element => {
        const key = element.getAttribute('data-i18n-key');
        const translation = translations[lang][key];
        if (translation) {
            if (element.hasAttribute('placeholder')) {
                element.placeholder = translation;
            } else {
                element.innerHTML = translation;
            }
        } else {
            console.warn(`Translation key "${key}" not found for language "${lang}".`);
        }
    });
}

function t(key) {
    return translations[currentLanguage][key] || key;
}
