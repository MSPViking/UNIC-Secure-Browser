Office.onReady(() => {
    // Only run this when the Office context is ready.
    if (Office.context.mailbox.item) {
        const item = Office.context.mailbox.item;
        getBody(item).then(body => {
            const urls = extractUrls(body);
            displayUrls(urls);
        }).catch(error => {
            console.error('Error:', error);
        });
    }
});

async function getBody(item) {
    return new Promise((resolve, reject) => {
        item.body.getAsync("text", (result) => {
            if (result.status === Office.AsyncResultStatus.Succeeded) {
                resolve(result.value);
            } else {
                reject(result.error);
            }
        });
    });
}

function extractUrls(text) {
    const urlRegex = /(https?:\/\/[^\s]+)/g;
    return text.match(urlRegex) || [];
}

function displayUrls(urls) {
    const container = document.getElementById('links-container');
    urls.forEach(url => {
        const linkElement = document.createElement('div');
        linkElement.innerHTML = `<a href="${url}" target="_blank">${url}</a>`;
        const button = document.createElement('button');
        button.textContent = 'Open in Secure Browser';
        button.onclick = () => openInSecureBrowser(url);
        linkElement.appendChild(button);
        container.appendChild(linkElement);
    });
}

async function openInSecureBrowser(url) {
    const response = await fetch('http://109.189.76.223:9900/process-url', {
        method: 'POST',
        headers: {
            'Content-Type': 'application/json',
        },
        body: JSON.stringify({ url: url }),
    });
    const result = await response.text();
    window.open(result, '_blank');
}
