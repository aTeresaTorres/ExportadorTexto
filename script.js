// Función auxiliar para mostrar mensajes
function showMessage(message, type) {
    const messageDiv = document.getElementById('message');
    messageDiv.textContent = message;
    messageDiv.className = `message ${type}`;
    setTimeout(() => {
        messageDiv.className = 'message';
    }, 4000);
}

// IMPORTAR ARCHIVO
function importFile() {
    const fileInput = document.getElementById('fileImport');
    const file = fileInput.files[0];
    
    if (!file) return;
    
    // Obtener el nombre del archivo sin extensión
    const filename = file.name.split('.').slice(0, -1).join('.');
    document.getElementById('filename').value = filename;
    
    // Leer el archivo
    const reader = new FileReader();
    
    reader.onload = function(e) {
        const fileContent = e.target.result;
        
        // Si es un archivo binario (DOCX, ODT), intentar extraer el texto
        if (file.type === 'application/vnd.openxmlformats-officedocument.wordprocessingml.document' || 
            file.name.endsWith('.docx')) {
            extractTextFromDocx(file, filename, fileContent);
        } else if (file.type === 'application/vnd.oasis.opendocument.text' || 
                   file.name.endsWith('.odt')) {
            extractTextFromOdt(file, filename, fileContent);
        } else {
            // Para otros formatos (TXT, RTF, HTML), mostrar el contenido directamente
            document.getElementById('textContent').value = fileContent;
            showMessage(`Archivo "${file.name}" importado correctamente`, 'success');
        }
    };
    
    reader.onerror = function() {
        showMessage('Error al leer el archivo', 'error');
    };
    
    // Leer como texto para la mayoría de formatos
    if (file.type === 'application/vnd.openxmlformats-officedocument.wordprocessingml.document' || 
        file.name.endsWith('.docx')) {
        reader.readAsArrayBuffer(file);
    } else if (file.type === 'application/vnd.oasis.opendocument.text' || 
               file.name.endsWith('.odt')) {
        reader.readAsArrayBuffer(file);
    } else {
        reader.readAsText(file);
    }
}

// Extraer texto de DOCX
async function extractTextFromDocx(file, filename, fileContent) {
    try {
        const zip = await JSZip.loadAsync(file);
        const xmlContent = await zip.file('word/document.xml').async('string');
        
        // Parsear XML y extraer texto
        const parser = new DOMParser();
        const xmlDoc = parser.parseFromString(xmlContent, 'text/xml');
        const textElements = xmlDoc.querySelectorAll('w\\:t, t');
        
        let text = '';
        textElements.forEach(element => {
            text += element.textContent;
        });
        
        if (text.trim()) {
            document.getElementById('textContent').value = text;
            showMessage(`Archivo DOCX "${file.name}" importado correctamente`, 'success');
        } else {
            showMessage(`No se pudo extraer texto del DOCX. Por favor, usa TXT o RTF`, 'error');
        }
    } catch (error) {
        showMessage(`Error al procesar DOCX. Por favor, usa TXT o RTF`, 'error');
        console.error(error);
    }
}

// Extraer texto de ODT
async function extractTextFromOdt(file, filename, fileContent) {
    try {
        const zip = await JSZip.loadAsync(file);
        const xmlContent = await zip.file('content.xml').async('string');
        
        // Parsear XML y extraer texto
        const parser = new DOMParser();
        const xmlDoc = parser.parseFromString(xmlContent, 'text/xml');
        const textElements = xmlDoc.querySelectorAll('text\\:p, p');
        
        let text = '';
        textElements.forEach((element, index) => {
            if (index > 0) text += '\n';
            text += element.textContent;
        });
        
        if (text.trim()) {
            document.getElementById('textContent').value = text;
            showMessage(`Archivo ODT "${file.name}" importado correctamente`, 'success');
        } else {
            showMessage(`No se pudo extraer texto del ODT. Por favor, usa TXT o RTF`, 'error');
        }
    } catch (error) {
        showMessage(`Error al procesar ODT. Por favor, usa TXT o RTF`, 'error');
        console.error(error);
    }
}

// Validar que hay texto
function validateText() {
    const text = document.getElementById('textContent').value.trim();
    const filename = document.getElementById('filename').value.trim();
    
    if (!text) {
        showMessage('Por favor escribe algo antes de exportar', 'error');
        return null;
    }
    
    if (!filename) {
        showMessage('Por favor asigna un nombre al archivo', 'error');
        return null;
    }
    
    return { text, filename };
}

// Descargar archivo
function downloadFile(content, filename, mimeType) {
    const blob = new Blob([content], { type: mimeType });
    const url = URL.createObjectURL(blob);
    const link = document.createElement('a');
    link.href = url;
    link.download = filename;
    document.body.appendChild(link);
    link.click();
    document.body.removeChild(link);
    URL.revokeObjectURL(url);
}

// 1. EXPORTAR A TXT
function exportTXT() {
    const data = validateText();
    if (!data) return;
    
    downloadFile(data.text, `${data.filename}.txt`, 'text/plain');
    showMessage(`Archivo ${data.filename}.txt descargado correctamente`, 'success');
}

// 2. EXPORTAR A DOCX
async function exportDOCX() {
    const data = validateText();
    if (!data) return;
    
    try {
        const doc = new docx.Document({
            sections: [{
                properties: {},
                children: [
                    new docx.Paragraph({
                        text: data.text,
                        spacing: {
                            line: 480,
                            lineRule: "auto"
                        }
                    })
                ]
            }]
        });

        const blob = await docx.Packer.toBlob(doc);
        const url = URL.createObjectURL(blob);
        const link = document.createElement('a');
        link.href = url;
        link.download = `${data.filename}.docx`;
        document.body.appendChild(link);
        link.click();
        document.body.removeChild(link);
        URL.revokeObjectURL(url);
        showMessage(`Archivo ${data.filename}.docx descargado correctamente`, 'success');
    } catch (error) {
        showMessage('Error al crear el archivo DOCX', 'error');
        console.error(error);
    }
}

// 3. EXPORTAR A PDF
function exportPDF() {
    const data = validateText();
    if (!data) return;
    
    try {
        const { jsPDF } = window.jspdf;
        const doc = new jsPDF();
        
        // Márgenes
        const marginLeft = 10;
        const marginTop = 10;
        const pageHeight = doc.internal.pageSize.getHeight();
        const pageWidth = doc.internal.pageSize.getWidth();
        const maxWidth = pageWidth - (marginLeft * 2);
        
        // Set font
        doc.setFont('helvetica', 'normal');
        doc.setFontSize(12);
        
        // Dividir el texto en líneas que caben en la página
        const lines = doc.splitTextToSize(data.text, maxWidth);
        
        let yPosition = marginTop;
        const lineHeight = 7;
        
        lines.forEach((line) => {
            if (yPosition > pageHeight - 10) {
                doc.addPage();
                yPosition = marginTop;
            }
            doc.text(line, marginLeft, yPosition);
            yPosition += lineHeight;
        });
        
        doc.save(`${data.filename}.pdf`);
        showMessage(`Archivo ${data.filename}.pdf descargado correctamente`, 'success');
    } catch (error) {
        showMessage('Error al crear el archivo PDF', 'error');
        console.error(error);
    }
}

// 4. EXPORTAR A RTF
function exportRTF() {
    const data = validateText();
    if (!data) return;
    
    try {
        // Escapar caracteres especiales RTF
        const escapedText = data.text
            .replace(/\\/g, '\\\\')
            .replace(/\{/g, '\\{')
            .replace(/\}/g, '\\}')
            .replace(/\n/g, '\\par\n');
        
        const rtf = `{\\rtf1\\ansi\\ansicpg1252\\deff0
{\\fonttbl{\\f0 Times New Roman;}}
{\\colortbl;\\red0\\green0\\blue0;}
\\uc1\\pard\\plain\\deftab720\\f0\\fs20
${escapedText}
}`;
        
        downloadFile(rtf, `${data.filename}.rtf`, 'application/rtf');
        showMessage(`Archivo ${data.filename}.rtf descargado correctamente`, 'success');
    } catch (error) {
        showMessage('Error al crear el archivo RTF', 'error');
        console.error(error);
    }
}

// 5. EXPORTAR A HTML
function exportHTML() {
    const data = validateText();
    if (!data) return;
    
    try {
        const htmlContent = `<!DOCTYPE html>
<html lang="es">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>${data.filename}</title>
    <style>
        body {
            font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif;
            max-width: 900px;
            margin: 0 auto;
            padding: 20px;
            line-height: 1.6;
            color: #333;
        }
        p {
            text-align: justify;
        }
    </style>
</head>
<body>
    <div>
        ${data.text.split('\n').map(line => `<p>${escapeHtml(line)}</p>`).join('\n        ')}
    </div>
</body>
</html>`;
        
        downloadFile(htmlContent, `${data.filename}.html`, 'text/html');
        showMessage(`Archivo ${data.filename}.html descargado correctamente`, 'success');
    } catch (error) {
        showMessage('Error al crear el archivo HTML', 'error');
        console.error(error);
    }
}

// 6. EXPORTAR A ODT
async function exportODT() {
    const data = validateText();
    if (!data) return;
    
    try {
        // ODT es un formato ZIP con XML
        const zip = new JSZip();
        
        // Agregar el archivo mimetype
        zip.file('mimetype', 'application/vnd.oasis.opendocument.text', { binary: false });
        
        // Crear la estructura de carpetas
        const metaInf = zip.folder('META-INF');
        const metaXml = `<?xml version="1.0" encoding="UTF-8"?>
<manifest:manifest xmlns:manifest="urn:oasis:names:tc:opendocument:xmlns:manifest:1.0">
    <manifest:file-entry manifest:media-type="application/vnd.oasis.opendocument.text" manifest:full-path="/"/>
    <manifest:file-entry manifest:media-type="text/xml" manifest:full-path="content.xml"/>
    <manifest:file-entry manifest:media-type="text/xml" manifest:full-path="styles.xml"/>
    <manifest:file-entry manifest:media-type="text/xml" manifest:full-path="meta.xml"/>
</manifest:manifest>`;
        metaInf.file('manifest.xml', metaXml);
        
        // Crear content.xml con el contenido
        const contentXml = `<?xml version="1.0" encoding="UTF-8"?>
<office:document-content 
    xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0"
    xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0"
    xmlns:style="urn:oasis:names:tc:opendocument:xmlns:style:1.0">
    <office:body>
        <office:text>
            ${data.text.split('\n').map(line => `<text:p>${escapeXml(line)}</text:p>`).join('\n            ')}
        </office:text>
    </office:body>
</office:document-content>`;
        zip.file('content.xml', contentXml);
        
        // Crear styles.xml
        const stylesXml = `<?xml version="1.0" encoding="UTF-8"?>
<office:document-styles xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0">
    <office:styles/>
</office:document-styles>`;
        zip.file('styles.xml', stylesXml);
        
        // Crear meta.xml
        const metaXmlFile = `<?xml version="1.0" encoding="UTF-8"?>
<office:document-meta xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0">
    <office:meta/>
</office:document-meta>`;
        zip.file('meta.xml', metaXmlFile);
        
        // Generar y descargar
        const blob = await zip.generateAsync({ type: 'blob' });
        const url = URL.createObjectURL(blob);
        const link = document.createElement('a');
        link.href = url;
        link.download = `${data.filename}.odt`;
        document.body.appendChild(link);
        link.click();
        document.body.removeChild(link);
        URL.revokeObjectURL(url);
        
        showMessage(`Archivo ${data.filename}.odt descargado correctamente`, 'success');
    } catch (error) {
        showMessage('Error al crear el archivo ODT', 'error');
        console.error(error);
    }
}

// Funciones auxiliares para escapar caracteres
function escapeHtml(text) {
    const map = {
        '&': '&amp;',
        '<': '&lt;',
        '>': '&gt;',
        '"': '&quot;',
        "'": '&#039;'
    };
    return text.replace(/[&<>"']/g, m => map[m]);
}

function escapeXml(text) {
    const map = {
        '&': '&amp;',
        '<': '&lt;',
        '>': '&gt;',
        '"': '&quot;',
        "'": '&apos;'
    };
    return text.replace(/[&<>"']/g, m => map[m]);
}

// 7. EXPORTAR TODOS LOS FORMATOS
async function exportAll() {
    const data = validateText();
    if (!data) return;
    
    try {
        // Hay que hacer esto secuencialmente para que no haya conflictos
        showMessage('Generando archivos...', 'success');
        
        // TXT
        downloadFile(data.text, `${data.filename}.txt`, 'text/plain');
        await new Promise(resolve => setTimeout(resolve, 300));
        
        // DOCX
        const docxDoc = new docx.Document({
            sections: [{
                properties: {},
                children: [
                    new docx.Paragraph({
                        text: data.text,
                        spacing: {
                            line: 480,
                            lineRule: "auto"
                        }
                    })
                ]
            }]
        });
        const docxBlob = await docx.Packer.toBlob(docxDoc);
        const docxUrl = URL.createObjectURL(docxBlob);
        const docxLink = document.createElement('a');
        docxLink.href = docxUrl;
        docxLink.download = `${data.filename}.docx`;
        document.body.appendChild(docxLink);
        docxLink.click();
        document.body.removeChild(docxLink);
        URL.revokeObjectURL(docxUrl);
        await new Promise(resolve => setTimeout(resolve, 300));
        
        // PDF
        const { jsPDF } = window.jspdf;
        const doc = new jsPDF();
        const marginLeft = 10;
        const marginTop = 10;
        const pageHeight = doc.internal.pageSize.getHeight();
        const pageWidth = doc.internal.pageSize.getWidth();
        const maxWidth = pageWidth - (marginLeft * 2);
        doc.setFont('helvetica', 'normal');
        doc.setFontSize(12);
        const lines = doc.splitTextToSize(data.text, maxWidth);
        let yPosition = marginTop;
        const lineHeight = 7;
        lines.forEach((line) => {
            if (yPosition > pageHeight - 10) {
                doc.addPage();
                yPosition = marginTop;
            }
            doc.text(line, marginLeft, yPosition);
            yPosition += lineHeight;
        });
        doc.save(`${data.filename}.pdf`);
        await new Promise(resolve => setTimeout(resolve, 300));
        
        // RTF
        const escapedText = data.text
            .replace(/\\/g, '\\\\')
            .replace(/\{/g, '\\{')
            .replace(/\}/g, '\\}')
            .replace(/\n/g, '\\par\n');
        const rtf = `{\\rtf1\\ansi\\ansicpg1252\\deff0
{\\fonttbl{\\f0 Times New Roman;}}
{\\colortbl;\\red0\\green0\\blue0;}
\\uc1\\pard\\plain\\deftab720\\f0\\fs20
${escapedText}
}`;
        downloadFile(rtf, `${data.filename}.rtf`, 'application/rtf');
        await new Promise(resolve => setTimeout(resolve, 300));
        
        // HTML
        const htmlContent = `<!DOCTYPE html>
<html lang="es">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>${data.filename}</title>
    <style>
        body {
            font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif;
            max-width: 900px;
            margin: 0 auto;
            padding: 20px;
            line-height: 1.6;
            color: #333;
        }
        p {
            text-align: justify;
        }
    </style>
</head>
<body>
    <div>
        ${data.text.split('\n').map(line => `<p>${escapeHtml(line)}</p>`).join('\n        ')}
    </div>
</body>
</html>`;
        downloadFile(htmlContent, `${data.filename}.html`, 'text/html');
        await new Promise(resolve => setTimeout(resolve, 300));
        
        // ODT
        const zip = new JSZip();
        zip.file('mimetype', 'application/vnd.oasis.opendocument.text', { binary: false });
        const metaInf = zip.folder('META-INF');
        const metaXml = `<?xml version="1.0" encoding="UTF-8"?>
<manifest:manifest xmlns:manifest="urn:oasis:names:tc:opendocument:xmlns:manifest:1.0">
    <manifest:file-entry manifest:media-type="application/vnd.oasis.opendocument.text" manifest:full-path="/"/>
    <manifest:file-entry manifest:media-type="text/xml" manifest:full-path="content.xml"/>
    <manifest:file-entry manifest:media-type="text/xml" manifest:full-path="styles.xml"/>
    <manifest:file-entry manifest:media-type="text/xml" manifest:full-path="meta.xml"/>
</manifest:manifest>`;
        metaInf.file('manifest.xml', metaXml);
        const contentXml = `<?xml version="1.0" encoding="UTF-8"?>
<office:document-content 
    xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0"
    xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0"
    xmlns:style="urn:oasis:names:tc:opendocument:xmlns:style:1.0">
    <office:body>
        <office:text>
            ${data.text.split('\n').map(line => `<text:p>${escapeXml(line)}</text:p>`).join('\n            ')}
        </office:text>
    </office:body>
</office:document-content>`;
        zip.file('content.xml', contentXml);
        const stylesXml = `<?xml version="1.0" encoding="UTF-8"?>
<office:document-styles xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0">
    <office:styles/>
</office:document-styles>`;
        zip.file('styles.xml', stylesXml);
        const metaXmlFile = `<?xml version="1.0" encoding="UTF-8"?>
<office:document-meta xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0">
    <office:meta/>
</office:document-meta>`;
        zip.file('meta.xml', metaXmlFile);
        const zipBlob = await zip.generateAsync({ type: 'blob' });
        const zipUrl = URL.createObjectURL(zipBlob);
        const zipLink = document.createElement('a');
        zipLink.href = zipUrl;
        zipLink.download = `${data.filename}.odt`;
        document.body.appendChild(zipLink);
        zipLink.click();
        document.body.removeChild(zipLink);
        URL.revokeObjectURL(zipUrl);
        
        showMessage(`Se han descargado todos los 6 formatos correctamente`, 'success');
    } catch (error) {
        showMessage('Error al crear los archivos', 'error');
        console.error(error);
    }
}
