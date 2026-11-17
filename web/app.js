import * as pdfjs from "https://cdn.jsdelivr.net/npm/pdfjs-dist@4.10.38/build/pdf.mjs";

pdfjs.GlobalWorkerOptions.workerSrc =
    "https://cdn.jsdelivr.net/npm/pdfjs-dist@4.10.38/build/pdf.worker.mjs";

const form = document.querySelector("#timesheet-form");
const input = document.querySelector("#pdf");
const button = document.querySelector("#generate");
const status = document.querySelector("#status");

let pyodide;

function setStatus(message) {
    status.textContent = message;
}

function positionedItems(items, viewportHeight) {
    return items
        .filter(item => item.str && item.str.trim())
        .map(item => ({
            text: item.str,
            x: item.transform[4],
            y: viewportHeight - item.transform[5] - Math.abs(item.height || 0),
            width: item.width,
            height: Math.abs(item.height || 0),
        }));
}

async function extractPages(file) {
    const data = new Uint8Array(await file.arrayBuffer());
    const document = await pdfjs.getDocument({ data }).promise;
    const pages = [];

    try {
        for (let number = 1; number <= document.numPages; number++) {
            setStatus(`Lendo página ${number} de ${document.numPages}...`);

            const page = await document.getPage(number);
            const content = await page.getTextContent();

            pages.push(positionedItems(content.items, page.getViewport({ scale: 1 }).height));
        }
    } finally {
        await document.destroy();
    }

    return pages;
}

async function initialize() {
    try {
        pyodide = await loadPyodide({
            indexURL: "https://cdn.jsdelivr.net/pyodide/v0.28.3/full/",
        });

        setStatus("Instalando biblioteca Excel...");

        await pyodide.loadPackage("micropip");
        await pyodide.runPythonAsync(
            "import micropip\nawait micropip.install('openpyxl==3.1.5')"
        );

        const names = ["extraction.py", "core.py", "excel_adapter.py", "labels.json", "bridge.py"];

        for (const name of names) {
            const url = name === "bridge.py" ? "./bridge.py" : `./python/${name}`;
            const response = await fetch(url);

            if (!response.ok) {
                throw new Error(`Não foi possível carregar ${name}`);
            }

            pyodide.FS.writeFile(name, await response.text());
        }

        await pyodide.runPythonAsync("from bridge import process_pdf_pages");

        button.disabled = false;
        button.textContent = "Gerar Excel";
        setStatus("Pronto.");
    } catch (error) {
        setStatus(`Erro ao iniciar: ${error.message}`);
        console.error(error);
    }
}

form.addEventListener("submit", async event => {
    event.preventDefault();

    const file = input.files[0];

    if (!file || !pyodide) {
        return;
    }

    button.disabled = true;

    try {
        const pages = await extractPages(file);

        setStatus("Calculando jornadas e gerando Excel...");

        pyodide.globals.set("uploaded_pages_json", JSON.stringify(pages));

        const path = await pyodide.runPythonAsync(
            "process_pdf_pages(uploaded_pages_json)"
        );

        const bytes = pyodide.FS.readFile(path);
        const blob = new Blob([bytes], {
            type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        });

        const url = URL.createObjectURL(blob);
        const link = document.createElement("a");

        link.href = url;
        link.download = "folha_de_ponto.xlsx";
        link.click();

        URL.revokeObjectURL(url);
        setStatus("Planilha gerada.");
    } catch (error) {
        setStatus(`Erro: ${error.message}`);
        console.error(error);
    } finally {
        button.disabled = false;
    }
});

initialize();
