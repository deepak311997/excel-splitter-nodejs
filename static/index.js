function onSubmit() {
    document.getElementById('submit').disabled = true;
    document.querySelector('.loading').style.display = "block";
    var xhr = new XMLHttpRequest();

    xhr.open("POST", "/api/process");
    xhr.responseType = 'blob';
    xhr.onload = function ({ target: { response, status }}) {
        document.getElementById('submit').disabled = false;
        document.querySelector('.loading').style.display = "none";
        const fileName = document.getElementById('output').value;
        
        if (status === 200) {
            var blob = new Blob([event.target.response], { type: 'application/zip' });
            var link = document.createElement('a');
            link.href = window.URL.createObjectURL(blob);
            link.download = fileName ? `${fileName}.zip` : 'result.zip';
            document.body.appendChild(link);
            link.click();
            document.body.removeChild(link);
        } else {
            alert(`${status}: Invalid payload`);
        }
    };
    var formData = new FormData(document.getElementById("excel-form"));
    xhr.send(formData);
}
window.onload = function() {
    document.getElementById('submit').addEventListener('click', onSubmit);
}