function showUploadOptions() {
    Swal.fire({
        title: 'Pilih Format',
        showDenyButton: true,
        showCancelButton: true,
        confirmButtonText: 'File Excel',
        denyButtonText: 'Data Text',
        cancelButtonText: 'Kembali',
        allowOutsideClick: false,
        customClass: {
            confirmButton: 'excel-button'
        }
    }).then((result) => {
        if (result.isConfirmed) {
            showExcelUploadDialog();
        } else if (result.isDenied) {
            showTextUploadDialog();
        }
    });
}

function showExcelUploadDialog(existingFile) {
    Swal.fire({
        title: 'Upload File Excel',
        input: 'file',
        inputAttributes: {
            'accept': '.xlsx, .xls, .csv',
            'aria-label': 'Upload your Excel file'
        },
        inputValue: existingFile || null,
        showCancelButton: true,
        confirmButtonText: 'Import Data',
        cancelButtonText: 'Kembali',
        allowOutsideClick: false,
        customClass: {
            confirmButton: 'uploadfile-button'
        }
    }).then((fileResult) => {
        if (fileResult.isDismissed) {
            showUploadOptions();
        } else if (fileResult.value) {
            checkFileFormat(fileResult.value, 'excel');
        } else {
            showWarning('Anda harus mengunggah file untuk melanjutkan.', showExcelUploadDialog);
        }
    });
}

function showTextUploadDialog(existingText) {
    Swal.fire({
        title: 'Enter Text Data',
        input: 'textarea',
        inputPlaceholder: 'Enter text data here...',
        inputValue: existingText || '',
        showCancelButton: true,
        confirmButtonText: 'Import Data',
        cancelButtonText: 'Kembali',
        allowOutsideClick: false,
        customClass: {
            confirmButton: 'uploadfile-button'
        }
    }).then((textResult) => {
        if (textResult.isDismissed) {
            showUploadOptions();
        } else if (textResult.value) {
            checkTextFormat(textResult.value, 'text');
        } else {
            showWarning('Anda harus memasukkan data teks untuk melanjutkan.', showTextUploadDialog);
        }
    });
}

function checkFileFormat(file, type) {
    const reader = new FileReader();
    reader.onload = function (e) {
        const data = new Uint8Array(e.target.result);
        const workbook = XLSX.read(data, { type: 'array' });
        const firstSheetName = workbook.SheetNames[0];
        const worksheet = workbook.Sheets[firstSheetName];
        const jsonData = XLSX.utils.sheet_to_json(worksheet, { header: 1 });

        if (jsonData.length > 0 && jsonData[0][0] === 'ID Agen' && jsonData[0][1] === 'Status') {
            showUploadConfirmation(() => uploadfileTableFromFile(file), type, file);
        } else {
            showWarning('Format file salah. Pastikan file mengandung kolom "ID Agen" dan "Status".', showExcelUploadDialog);
        }
    };
    reader.readAsArrayBuffer(file);
}

function checkTextFormat(textData, type) {
    const lines = textData.split('\n');
    if (lines.length > 0 && lines[0].trim().split(/\s+/)[0] === 'ID' && lines[0].trim().split(/\s+/)[1] === 'Agen') {
        showUploadConfirmation(() => uploadfileTableFromText(textData), type, textData);
    } else {
        showWarning('Format data yang diinput salah, Silahkan coba lagi dengan format yang benar!', showTextUploadDialog);
    }
}

function showUploadConfirmation(callback, type, inputData) {
    Swal.fire({
        title: 'Konfirmasi',
        html: '<div style="font-size: 15px; font-family: Calibri;">Transaksi hanya akan dihitung jika Status Transaksi adalah "SUCCESS". Data dengan status lain, seperti "Failed", tidak akan dihitung karena Tools ini hanya menghitung berdasarkan Status Transaksi "SUCCESS".<br><br>Jika Anda sudah yakin, silakan klik "Lanjutkan".</div>',
        icon: 'info',
        showCancelButton: true,
        confirmButtonText: 'Lanjutkan',
        cancelButtonText: 'Batal',
        allowOutsideClick: false,
        customClass: {
            confirmButton: 'uploadfile-button',
            cancelButton: 'swal2-cancel swal2-styled'
        }
    }).then((result) => {
        if (result.isConfirmed) {
            callback();
        } else if (result.isDismissed) {
            if (type === 'excel') {
                showExcelUploadDialog(inputData);
            } else if (type === 'text') {
                showTextUploadDialog(inputData);
            }
        }
    });
}

function uploadfileTableFromFile(file) {
    Swal.fire({
        title: 'Processing...',
        text: 'Mohon ditunggu sedang proses import data.',
        allowOutsideClick: false,
        didOpen: () => {
            Swal.showLoading();
        }
    });

    const reader = new FileReader();
    reader.onload = function (e) {
        setTimeout(() => {
            const data = new Uint8Array(e.target.result);
            const workbook = XLSX.read(data, { type: 'array' });
            const firstSheetName = workbook.SheetNames[0];
            const worksheet = workbook.Sheets[firstSheetName];
            const jsonData = XLSX.utils.sheet_to_json(worksheet, { header: 1 });

            if (jsonData.length > 1) {
                const transactions = {};
                jsonData.slice(1).forEach(row => {
                    const id = row[0];
                    const status = row[1];
                    if (status === 'SUCCESS') {
                        if (transactions[id]) {
                            transactions[id]++;
                        } else {
                            transactions[id] = 1;
                        }
                    }
                });

                const sortedTransactions = Object.keys(transactions).map(id => ({
                    id,
                    count: transactions[id]
                })).sort((a, b) => b.count - a.count);

                const resultTable = document.getElementById('resultTable').getElementsByTagName('tbody')[0];
                resultTable.innerHTML = '';
                sortedTransactions.forEach((item, index) => {
                    const row = resultTable.insertRow();
                    row.insertCell(0).innerText = index + 1;
                    row.insertCell(1).innerText = item.id;
                    row.insertCell(2).innerText = formatNumber(item.count);
                    row.insertCell(3).innerText = 'Level ' + (index + 1);
                });

                updateSummary(sortedTransactions.length, sortedTransactions.reduce((sum, item) => sum + item.count, 0));

                Swal.fire('Success', 'Data Berhasil Terkirim:)', 'success');
                showExportResetButtons();
                document.getElementById('summary').style.display = 'block';
            } else {
                showWarning('File tidak mengandung data yang valid.', showExcelUploadDialog);
            }
        }, 2000);
    };

    reader.readAsArrayBuffer(file);
}

function uploadfileTableFromText(textData) {
    Swal.fire({
        title: 'Processing...',
        text: 'Mohon ditunggu sedang proses import data.',
        allowOutsideClick: false,
        didOpen: () => {
            Swal.showLoading();
        }
    });

    setTimeout(() => {
        const transactions = {};
        const lines = textData.split('\n');

        if (lines.length > 0 && lines[0].trim() !== "") {
            lines.slice(1).forEach(line => {
                const [id, status] = line.trim().split(/\s+/);
                if (status === 'SUCCESS') {
                    if (transactions[id]) {
                        transactions[id]++;
                    } else {
                        transactions[id] = 1;
                    }
                }
            });

            const sortedTransactions = Object.keys(transactions).map(id => ({
                id,
                count: transactions[id]
            })).sort((a, b) => b.count - a.count);

            const resultTable = document.getElementById('resultTable').getElementsByTagName('tbody')[0];
            resultTable.innerHTML = '';
            sortedTransactions.forEach((item, index) => {
                const row = resultTable.insertRow();
                row.insertCell(0).innerText = index + 1;
                row.insertCell(1).innerText = item.id;
                row.insertCell(2).innerText = formatNumber(item.count);
                row.insertCell(3).innerText = 'Level ' + (index + 1);
            });

            updateSummary(sortedTransactions.length, sortedTransactions.reduce((sum, item) => sum + item.count, 0));

            Swal.fire('Success', 'Data berhasil diupload.', 'success');
            showExportResetButtons();
            document.getElementById('summary').style.display = 'block';
        } else {
            showWarning('Data teks tidak mengandung informasi yang valid.', showTextUploadDialog);
        }
    }, 2000);
}

function updateSummary(jumlahagen, jumlatrxhagen) {
    const summary = document.getElementById('summary');
    summary.innerHTML = `Jumlah Agen: ${jumlahagen} <br> Jumlah Transaksi: ${formatNumber(jumlatrxhagen)}`;
}

function showExportResetButtons() {
    const uploadContainer = document.getElementById('uploadContainer');
    uploadContainer.innerHTML = `
<button class="upload-container" style="background: green;" onclick="exportData()">Export Data</button>
<button class="upload-container" onclick="resetData()">Reset Data</button>
`;
}

function exportData() {
    Swal.fire({
        title: 'Export Data',
        html: '<div style="font-size: 17px; text-align: center;">Tuliskan nama file sesuai keinginanmu.</div>',
        input: 'text',
        inputPlaceholder: 'Tuliskan nama file...',
        showCancelButton: true,
        confirmButtonText: 'Lanjutkan',
        cancelButtonText: 'Batal',
        allowOutsideClick: false,
        customClass: {
            confirmButton: 'uploadfile-button',
            cancelButton: 'swal2-cancel swal2-styled'
        }
    }).then((result) => {
        if (result.dismiss === Swal.DismissReason.cancel) {
            return; // Jika tombol Batal diklik, keluar dari fungsi tanpa melakukan apapun
        }

        if (!result.value) {
            Swal.fire({
                title: 'Peringatan',
                html: 'Harap masukkan nama file. Misalnya "Balapan transaksi bulan Juni 2024" silahkan bisa diisi sesuai keinginan.',
                icon: 'warning',
                confirmButtonText: 'Ok',
                allowOutsideClick: false,
                customClass: {
                    confirmButton: 'uploadfile-button'
                }
            }).then(() => {
                exportData(); // Kembali ke popup export data jika nama file tidak diisi
            });
        } else {
            const fileName = result.value;
            Swal.fire({
                title: 'Exporting...',
                text: 'Harap ditunggu, sedang proses export data.',
                allowOutsideClick: false,
                didOpen: () => {
                    Swal.showLoading();
                }
            });

            setTimeout(() => {
                try {
                    exportToExcel(fileName);
                    Swal.fire('Success', 'Data berhasil diexport.', 'success');
                } catch (error) {
                    Swal.fire('Error', 'Gagal mengekspor data.', 'error');
                }
            }, 2000);
        }
    });
}

function exportToExcel(fileName) {
    const table = document.getElementById('resultTable');
    const wb = XLSX.utils.book_new();
    const ws = XLSX.utils.table_to_sheet(table);

    const sheetName = 'Balapan Transaksi';
    XLSX.utils.book_append_sheet(wb, ws, sheetName);

    const range = XLSX.utils.decode_range(ws['!ref']);
    for (let R = range.s.r; R <= range.e.r; ++R) {
        for (let C = range.s.c; C <= range.e.c; ++C) {
            const cell_address = XLSX.utils.encode_cell({ r: R, c: C });
            if (!ws[cell_address]) ws[cell_address] = {};
            if (!ws[cell_address].s) ws[cell_address].s = {};
            ws[cell_address].s.border = {
                top: { style: "thin", color: { auto: 1 } },
                bottom: { style: "thin", color: { auto: 1 } },
                left: { style: "thin", color: { auto: 1 } },
                right: { style: "thin", color: { auto: 1 } }
            };
            ws[cell_address].s.alignment = { vertical: "center", horizontal: "center" };
            ws[cell_address].s.font = { name: "Calibri", sz: 11 };
        }
    }

    const headerRange = XLSX.utils.decode_range(ws['!ref']);
    for (let C = headerRange.s.c; C <= headerRange.e.c; ++C) {
        const cell_address = XLSX.utils.encode_cell({ r: 0, c: C });
        if (!ws[cell_address]) ws[cell_address] = {};
        if (!ws[cell_address].s) ws[cell_address].s = {};
        ws[cell_address].s.font = { bold: true, name: "Calibri", sz: 11 };
    }

    XLSX.writeFile(wb, `${fileName}.xlsx`);
}

function resetData() {
    const resultTable = document.getElementById('resultTable').getElementsByTagName('tbody')[0];
    resultTable.innerHTML = '';
    const uploadContainer = document.getElementById('uploadContainer');
    uploadContainer.innerHTML = `
<button id="uploadButton" onclick="showUploadOptions()">Upload Data</button>
<button id="tutorialButton" onclick="showCustomTutorial()">Tutorial</button>
`;
    const summary = document.getElementById('summary');
    summary.style.display = 'none';
    summary.innerHTML = 'Jumlah Agen: <br> Jumlah Transaksi:';

    Swal.fire({
        title: 'Success',
        text: 'Data berhasil direset.',
        icon: 'success',
        confirmButtonText: 'Ok',
        allowOutsideClick: false,
        customClass: {
            confirmButton: 'uploadfile-button'
        }
    });
}

function showWarning(message, retryCallback) {
    Swal.fire({
        title: 'Peringatan',
        text: message,
        icon: 'warning',
        confirmButtonText: 'Ok',
        allowOutsideClick: false,
        customClass: {
            confirmButton: 'uploadfile-button'
        }
    }).then(() => {
        retryCallback();
    });
}

function formatNumber(num) {
    return num.toString().replace(/\B(?=(\d{3})+(?!\d))/g, ".");
}

function showCustomTutorial() {
    const overlay = document.getElementById('customPopupOverlay');
    const videoFrame = document.getElementById('tutorialVideo');
    const loadingSpinner = document.getElementById('loadingSpinner');

    // Show the overlay
    overlay.style.display = 'flex';

    // Ensure the video plays properly by reloading the iframe source
    videoFrame.src += "&autoplay=1";

    // Show the spinner until the video is loaded
    videoFrame.onload = () => {
        loadingSpinner.style.display = 'none';
        videoFrame.style.display = 'block';
    };
}

function closeCustomPopup() {
    const overlay = document.getElementById('customPopupOverlay');
    const videoFrame = document.getElementById('tutorialVideo');

    // Hide the overlay
    overlay.style.display = 'none';

    // Stop the video and reset the iframe source
    videoFrame.src = videoFrame.src.replace("&autoplay=1", "");
    videoFrame.style.display = 'none';
    const loadingSpinner = document.getElementById('loadingSpinner');
    loadingSpinner.style.display = 'flex';
}

const style = document.createElement('style');
style.innerHTML = `
@keyframes spin {
0% { transform: rotate(0deg); }
100% { transform: rotate(360deg); }
}
`;
document.head.appendChild(style);