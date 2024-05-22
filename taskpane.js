Office.onReady(() => {
    if (Office.context.mailbox.item) {
        // When the add-in is ready, process the email body
        processEmailBody();
    }
});

function processEmailBody() {
    const item = Office.context.mailbox.item;
    item.body.getAsync("text", (result) => {
        if (result.status === Office.AsyncResultStatus.Succeeded) {
            const body = result.value;
            const urls = extractUrls(body);
            displayUrls(urls);
        } else {
            console.error(result.error);
        }
    });
}

function extractUrls(text) {
    const urlRegex = /(https?:\/\/[^\s]+)/g;
    return text.match(urlRegex) || [];
}

function displayUrls(urls) {
    const urlList = document.getElementById('url-list');
    urlList.innerHTML = ''; // Clear existing list

    urls.forEach(url => {
        const urlElement = document.createElement('div');
        urlElement.className = 'url-item';

        const link = document.createElement('a');
        link.href = url;
        link.textContent = url;
        link.target = '_blank';

        const icon = document.createElement('img');
        icon.src = 'https://github.com/MSPViking/UNIC-Secure-Browser/main/OIP.jpg'; // Your icon URL
        icon.alt = 'Open with UNIC Secure Browser';
        icon.className = 'secure-browser-icon';
        icon.addEventListener('click', () => openInSecureBrowser(url));

        urlElement.appendChild(link);
        urlElement.appendChild(icon);
        urlList.appendChild(urlElement);
    });
}

function openInSecureBrowser(url) {
    fetch('http://109.189.76.223:9900/process-url', {
        method: 'POST',
        headers: {
            'Content-Type': 'application/json',
        },
        body: JSON.stringify({ url }),
    })
    .then(response => response.text())
    .then(sessionUrl => {
        window.open(sessionUrl, '_blank');
    })
    .catch(error => {
        console.error('Error opening secure browser:', error);
    });
}
