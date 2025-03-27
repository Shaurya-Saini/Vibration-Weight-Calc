let weights = [];

document.getElementById('weight-form').addEventListener('submit', function (e) {
    e.preventDefault();

    const fileInput = document.getElementById('file-upload');
    const requiredWeight = parseFloat(document.getElementById('required-weight').value);
    const loadingDiv = document.getElementById('loading');

    if (!fileInput.files.length) {
        alert('Please upload an Excel file.');
        return;
    }

    if (isNaN(requiredWeight)) {
        alert('Please enter a valid required weight.');
        return;
    }

    const file = fileInput.files[0];
    const reader = new FileReader();

    loadingDiv.style.display = 'block';

    reader.onload = function (event) {
        const data = new Uint8Array(event.target.result);
        const workbook = XLSX.read(data, { type: 'array' });
        const sheetName = workbook.SheetNames[0];
        const worksheet = workbook.Sheets[sheetName];
        const json = XLSX.utils.sheet_to_json(worksheet, { header: 1 });

        weights = json.map(row => ({
            index: row[0],
            weight: parseFloat(row[1])
        })).filter(row => !isNaN(row.weight));

        if (weights.length === 0) {
            alert('No valid weight data found in the Excel file.');
            loadingDiv.style.display = 'none';
            return;
        }

        const closestPairs = findClosestWeightPairs(weights, requiredWeight);

        displayResults(closestPairs);
        showResultsPopup();
        loadingDiv.style.display = 'none';
    };

    reader.onerror = function (error) {
        alert('Error reading the Excel file.');
        console.error(error);
        loadingDiv.style.display = 'none';
    };

    reader.readAsArrayBuffer(file);
});

document.getElementById('show-weights-btn').addEventListener('click', function () {
    displayWeights();
    showWeightsPopup();
});

document.getElementById('excel-format-btn').addEventListener('click', function () {
    showExcelFormatPopup();
});

document.getElementById('download-results-btn').addEventListener('click', function () {
    downloadResults();
});

function findClosestWeightPairs(weights, requiredWeight) {
    let pairs = [];

    for (let i = 0; i < weights.length; i++) {
        for (let j = i + 1; j < weights.length; j++) {
            const sum = weights[i].weight + weights[j].weight;
            const difference = Math.abs(sum - requiredWeight);

            if (difference <= 0.5) {
                pairs.push({
                    index1: weights[i].index,
                    weight1: weights[i].weight,
                    index2: weights[j].index,
                    weight2: weights[j].weight,
                    difference: difference
                });
            }
        }
    }

    pairs = pairs.sort((a, b) => a.difference - b.difference);
    const weightCount = {};
    const filteredPairs = [];

    for (const pair of pairs) {
        const key1 = pair.index1;
        const key2 = pair.index2;

        if (!weightCount[key1]) weightCount[key1] = 0;
        if (!weightCount[key2]) weightCount[key2] = 0;

        if (weightCount[key1] < 3 && weightCount[key2] < 3) {
            filteredPairs.push(pair);
            weightCount[key1]++;
            weightCount[key2]++;
        }

        if (filteredPairs.length >= 10) break;
    }

    return filteredPairs;
}

function displayResults(pairs) {
    const resultsDiv = document.getElementById('results-list');
    resultsDiv.innerHTML = '';

    if (pairs.length) {
        pairs.forEach(pair => {
            resultsDiv.innerHTML += `
                <p>Index ${pair.index1}, Weight: ${pair.weight1.toFixed(1)} grams</p>
                <p>Index ${pair.index2}, Weight: ${pair.weight2.toFixed(1)} grams</p>
                <p>Total Weight: ${(pair.weight1 + pair.weight2).toFixed(2)} grams (Difference: ${pair.difference.toFixed(2)})</p>
                <hr>
            `;
        });
    } else {
        resultsDiv.innerHTML = '<p>No suitable pairs found.</p>';
    }
}

function downloadResults() {
    const resultsDiv = document.getElementById('results-list');
    const resultsText = resultsDiv.innerText;
    const blob = new Blob([resultsText], { type: 'text/plain' });
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = 'weight_calculator_results.txt';
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
    URL.revokeObjectURL(url);
}

function displayWeights() {
    const weightsDiv = document.getElementById('weights-list');
    weightsDiv.innerHTML = '';

    weights.forEach(weight => {
        weightsDiv.innerHTML += `<p>Index: ${weight.index}, Weight: ${weight.weight.toFixed(1)} grams</p>`;
    });
}

function showWeightsPopup() {
    const popup = document.getElementById('weights-popup');
    popup.style.display = 'block';
}

function showResultsPopup() {
    const popup = document.getElementById('results-popup');
    popup.style.display = 'block';
}

function showExcelFormatPopup() {
    const popup = document.getElementById('excel-format-popup');
    popup.style.display = 'block';
}

document.querySelectorAll('.close').forEach(btn => {
    btn.addEventListener('click', function () {
        const popup = this.closest('.popup');
        popup.style.display = 'none';
    });
});

window.onclick = function (event) {
    const weightsPopup = document.getElementById('weights-popup');
    const resultsPopup = document.getElementById('results-popup');
    const excelFormatPopup = document.getElementById('excel-format-popup');

    if (event.target == weightsPopup) {
        weightsPopup.style.display = 'none';
    }

    if (event.target == resultsPopup) {
        resultsPopup.style.display = 'none';
    }

    if (event.target == excelFormatPopup) {
        excelFormatPopup.style.display = 'none';
    }
}