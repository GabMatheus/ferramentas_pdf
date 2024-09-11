function handleFormSubmit(formId, actionUrl) {
    document.getElementById(formId).addEventListener('submit', function(event) {
        event.preventDefault();
        let formData = new FormData(this);

        fetch(actionUrl, {
            method: 'POST',
            body: formData
        })
        .then(response => response.json())
        .then(data => {
            if (data.message) {
                alert(data.message);
                if (data.output) {
                    // window.location.href = data.output;
                }
            } else {
                alert('Operação realizada com sucesso!');
            }
        })
        .catch(error => console.error('Error:', error));
    });
}

handleFormSubmit('rarForm', '/create_archive');
handleFormSubmit('docxForm', '/convert_docx');
handleFormSubmit('pdfToDocxForm', '/convert_pdf');
handleFormSubmit('mergePdfsForm', '/merge_pdfs');
