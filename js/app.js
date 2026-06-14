// --- State Global Aplikasi ---
let masterProduk = [];
let masterKasir = [];
let keranjang = [];
let daftarTransaksi = [];
let nomorRegistrasiSekarang = 100001; // Mulai dari No Reg ini

// Inisialisasi saat aplikasi pertama dibuka
document.addEventListener("DOMContentLoaded", () => {
    updateNoRegistrasi();
    setupEventListeners();
});

function updateNoRegistrasi() {
    document.getElementById("reg-number").value = `TRX-${nomorRegistrasiSekarang}`;
}

function setupEventListeners() {
    // Event Handler Input Excel/CSV
    document.getElementById("upload-produk").addEventListener("change", (e) => handleFileUpload(e, "produk"));
    document.getElementById("upload-user").addEventListener("change", (e) => handleFileUpload(e, "user"));
    
    // Event Operasional POS
    document.getElementById("btn-add-item").addEventListener("click", tambahItemKeKeranjang);
    document.getElementById("search-product").addEventListener("keypress", (e) => {
        if(e.key === 'Enter') tambahItemKeKeranjang();
    });
    
    document.getElementById("cash-payment").addEventListener("input", hitungKembalian);
    document.getElementById("btn-checkout").addEventListener("click", prosesCheckout);
    document.getElementById("btn-export-excel").addEventListener("click", eksporLaporanKeExcel);
}

// --- Fungsi Membaca File Excel / CSV ---
function handleFileUpload(event, tipe) {
    const file = event.target.files[0];
    if (!file) return;

    const reader = new FileReader();
    reader.onload = function(e) {
        const data = new Uint8Array(e.target.result);
        const workbook = XLSX.read(data, { type: 'array' });
        const sheetName = workbook.SheetNames[0];
        const sheet = workbook.Sheets[sheetName];
        const jsonData = XLSX.utils.sheet_to_json(sheet);

        if (tipe === "produk") {
            masterProduk = jsonData;
            const status = document.getElementById("status-produk");
            status.textContent = `Berhasil memuat ${masterProduk.length} Produk`;
            status.className = "status-msg success";
        } else if (tipe === "user") {
            masterKasir = jsonData;
            const status = document.getElementById("status-user");
            status.textContent = `Berhasil memuat ${masterKasir.length} Kasir`;
            status.className = "status-msg success";
            populateKasirDropdown();
        }
    };
    reader.readAsArrayBuffer(file);
}

// Mengisi pilihan kasir berdasarkan excel yang diupload
function populateKasirDropdown() {
    const select = document.getElementById("cashier-select");
    select.innerHTML = '<option value="">-- Pilih Kasir --</option>';
    
    masterKasir.forEach(user => {
        // Menyesuaikan nama kolom berdasarkan dataset (contoh: nama_member/kd_ksr/nama)
        const namaUser = user.nama_member || user.nama || user.kd_ksr || Object.values(user)[1];
        const kodeUser = user.kode_member || user.kd_ksr || Object.values(user)[0];
        
        if(namaUser) {
            const opt = document.createElement("option");
            opt.value = kodeUser;
            opt.textContent = `${kodeUser} - ${namaUser}`;
            select.appendChild(opt);
        }
    });
}

// --- Logika Keranjang Belanja ---
function tambahItemKeKeranjang() {
    const inputKeyword = document.getElementById("search-product").value.trim();
    if (!inputKeyword) return;

    // Cari berdasarkan 'plu' atau 'barcode' dari sheet Excel produk yang diupload
    const produk = masterProduk.find(p => String(p.plu) === inputKeyword || String(p.barcode) === inputKeyword);

    if (!produk) {
        alert("Produk tidak ditemukan! Periksa kembali kode PLU/Barcode atau pastikan Master Data sudah diupload.");
        return;
    }

    // Ambil data harga dan potongan harga (berdasarkan format database Anda)
    const hargaSatuan = parseFloat(produk.price1) || 0;
    const diskonRp = parseFloat(produk.disc1) || 0; // Mengakomodasi kolom hemat/diskon

    // Cek apakah barang sudah ada di keranjang
    const itemEksis = keranjang.find(item => item.plu === produk.plu);

    if (itemEksis) {
        itemEksis.qty += 1;
        itemEksis.total = itemEksis.qty * (itemEksis.hargaSatuan - itemEksis.diskonRp);
    } else {
        keranjang.push({
            plu: produk.plu,
            descp: produk.descp,
            hargaSatuan: hargaSatuan,
            qty: 1,
            diskonRp: diskonRp,
            total: 1 * (hargaSatuan - diskonRp)
        });
    }

    document.getElementById("search-product").value = ""; // reset input
    renderKeranjang();
}

function renderKeranjang() {
    const tbody = document.querySelector("#cart-table tbody");
    tbody.innerHTML = "";
    let grandTotal = 0;

    keranjang.forEach((item, index) => {
        grandTotal += item.total;
        const tr = document.createElement("tr");
        tr.innerHTML = `
            <td>${item.plu}</td>
            <td>${item.descp}</td>
            <td>Rp ${item.hargaSatuan.toLocaleString('id-ID')}</td>
            <td>${item.qty}</td>
            <td>Rp ${item.diskonRp.toLocaleString('id-ID')}</td>
            <td>Rp ${item.total.toLocaleString('id-ID')}</td>
            <td><button class="btn btn-danger" onclick="hapusItem(${index})">X</button></td>
        `;
        tbody.appendChild(tr);
    });

    document.getElementById("grand-total").textContent = `Rp ${grandTotal.toLocaleString('id-ID')}`;
    hitungKembalian();
}

function hapusItem(index) {
    keranjang.splice(index, 1);
    renderKeranjang();
}

function hitungKembalian() {
    const grandTotal = parseFloat(document.getElementById("grand-total").textContent.replace(/[^0-9]/g, '')) || 0;
    const cashPayment = parseFloat(document.getElementById("cash-payment").value) || 0;
    const kembalian = cashPayment - grandTotal;

    document.getElementById("cash-return").textContent = kembalian >= 0 ? `Rp ${kembalian.toLocaleString('id-ID')}` : "Rp 0";
}

// --- Proses Checkout & Cetak Struk ---
function prosesCheckout() {
    const kasirSelect = document.getElementById("cashier-select");
    if (!kasirSelect.value) {
        alert("Pilih Kasir aktif terlebih dahulu sebelum bertransaksi!");
        return;
    }
    if (keranjang.length === 0) {
        alert("Keranjang belanja masih kosong!");
        return;
    }

    const grandTotal = parseFloat(document.getElementById("grand-total").textContent.replace(/[^0-9]/g, '')) || 0;
    const cashPayment = parseFloat(document.getElementById("cash-payment").value) || 0;

    if (cashPayment < grandTotal) {
        alert("Pembayaran tunai kurang!");
        return;
    }

    const waktuSekarang = new Date();
    const stringWaktu = waktuSekarang.toLocaleDateString('id-ID') + ' ' + waktuSekarang.toLocaleTimeString('id-ID');
    const namaKasir = kasirSelect.options[kasirSelect.selectedIndex].text;
    const noReg = document.getElementById("reg-number").value;

    // 1. Simpan Ke Laporan Harian
    daftarTransaksi.push({
        noReg: noReg,
        waktu: stringWaktu,
        kasir: namaKasir,
        total: grandTotal
    });

    // 2. Buat Struk Kasir Thermal (HTML Print Only)
    buatStrukThermal(noReg, stringWaktu, namaKasir, grandTotal, cashPayment);

    // 3. Trigger Print Window (Otomatis memotong/mencetak pada printer thermal)
    window.print();

    // 4. Update UI untuk Transaksi Berikutnya
    renderLaporan();
    keranjang = [];
    renderKeranjang();
    document.getElementById("cash-payment").value = "";
    nomorRegistrasiSekarang++;
    updateNoRegistrasi();
}

function buatStrukThermal(noReg, waktu, kasir, total, bayar) {
    const containerStruk = document.getElementById("thermal-receipt");
    let itemsHtml = "";
    
    keranjang.forEach(item => {
        itemsHtml += `
            <div class="receipt-item">
                <span>${item.descp} (x${item.qty})</span>
                <span>${item.total.toLocaleString('id-ID')}</span>
            </div>
            ${item.diskonRp > 0 ? `<div class="receipt-item" style="font-size:10px; padding-left:10px;"><i>Hemat: -${(item.diskonRp * item.qty).toLocaleString('id-ID')}</i></div>` : ''}
        `;
    });

    containerStruk.innerHTML = `
        <div class="receipt-header">
            <h4>RETAIL MART</h4>
            <p>Jl. Jenderal Sudirman No. 12</p>
            <div class="receipt-line"></div>
            <p>No: ${noReg}<br>Tgl: ${waktu}<br>Kasir: ${kasir}</p>
        </div>
        <div class="receipt-line"></div>
        <div class="receipt-body">
            ${itemsHtml}
        </div>
        <div class="receipt-line"></div>
        <div class="receipt-total">
            <div class="receipt-item"><span>TOTAL:</span><span>Rp ${total.toLocaleString('id-ID')}</span></div>
            <div class="receipt-item"><span>BAYAR:</span><span>Rp ${bayar.toLocaleString('id-ID')}</span></div>
            <div class="receipt-item"><span>KEMBALI:</span><span>Rp ${(bayar - total).toLocaleString('id-ID')}</span></div>
        </div>
        <div class="receipt-line"></div>
        <div class="receipt-footer">
            <p>Terima Kasih<br>Selamat Berbelanja Kembali</p>
        </div>
    `;
}

function renderLaporan() {
    const tbody = document.querySelector("#report-table tbody");
    tbody.innerHTML = "";
    
    daftarTransaksi.forEach(trx => {
        const tr = document.createElement("tr");
        tr.innerHTML = `
            <td>${trx.noReg}</td>
            <td>${trx.waktu}</td>
            <td>${trx.kasir}</td>
            <td>Rp ${trx.total.toLocaleString('id-ID')}</td>
        `;
        tbody.appendChild(tr);
    });
}

// --- Fungsi Ekspor Laporan Harian ke Excel ---
function eksporLaporanKeExcel() {
    if (daftarTransaksi.length === 0) {
        alert("Belum ada data transaksi untuk diekspor hari ini.");
        return;
    }

    // Format data agar rapi dan profesional di Excel
    const dataSiapEkspor = daftarTransaksi.map(trx => ({
        "No. Registrasi": trx.noReg,
        "Tanggal & Jam Transaksi": trx.waktu,
        "Nama User / Kasir": trx.kasir,
        "Total Omset Penjualan (IDR)": trx.total
    }));

    // Membuat Worksheet & Workbook dengan SheetJS
    const worksheet = XLSX.utils.json_to_sheet(dataSiapEkspor);
    const workbook = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(workbook, worksheet, "Laporan Penjualan Harian");

    // Unduh File Excel secara otomatis ke komputer pengguna
    const tglHariIni = new Date().toISOString().split('T')[0];
    XLSX.writeFile(workbook, `LAPORAN_KASIR_${tglHariIni}.xlsx`);
}
