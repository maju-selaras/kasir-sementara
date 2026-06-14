// --- State Global ---
let masterProduk = [];
let masterKasir = [];
let keranjang = [];
let laporanHarian = []; // Untuk Sheet 1 (Laporan Penjualan)
let laporanDetail = []; // Untuk Sheet 2 (Rincian Produk)
let nomorRegistrasiHarian = 1;

document.addEventListener("DOMContentLoaded", () => {
    updateNoRegistrasi();
    
    // Event Listeners Upload Data
    document.getElementById("upload-produk").addEventListener("change", (e) => prosesUploadExcel(e, "produk"));
    document.getElementById("upload-user").addEventListener("change", (e) => prosesUploadExcel(e, "user"));
    
    // Event Operasional
    document.getElementById("btn-add-item").addEventListener("click", tambahKeKeranjang);
    document.getElementById("search-product").addEventListener("keypress", (e) => {
        if(e.key === 'Enter') tambahKeKeranjang();
    });
    
    document.getElementById("cash-payment").addEventListener("input", hitungKembalian);
    document.getElementById("btn-checkout").addEventListener("click", selesaikanTransaksi);
    document.getElementById("btn-export-excel").addEventListener("click", eksporLaporanKeExcel);
});

function updateNoRegistrasi() {
    const tanggalPrefix = new Date().toISOString().slice(0,10).replace(/-/g,"");
    document.getElementById("reg-number").value = `REG-${tanggalPrefix}-${String(nomorRegistrasiHarian).padStart(4, '0')}`;
}

// --- Membaca Excel ---
function prosesUploadExcel(event, jenis) {
    const file = event.target.files[0];
    if (!file) return;

    const reader = new FileReader();
    reader.onload = function(e) {
        const data = new Uint8Array(e.target.result);
        const workbook = XLSX.read(data, { type: 'array' });
        const sheet = workbook.Sheets[workbook.SheetNames[0]];
        const jsonData = XLSX.utils.sheet_to_json(sheet);

        if (jenis === "produk") {
            masterProduk = jsonData;
            const status = document.getElementById("status-produk");
            status.textContent = `Berhasil memuat ${masterProduk.length} Produk`;
            status.className = "status-msg success";
        } else if (jenis === "user") {
            masterKasir = jsonData;
            const status = document.getElementById("status-user");
            status.textContent = `Berhasil memuat ${masterKasir.length} Data User`;
            status.className = "status-msg success";
            isiDropdownKasir();
        }
    };
    reader.readAsArrayBuffer(file);
}

function isiDropdownKasir() {
    const select = document.getElementById("cashier-select");
    select.innerHTML = '<option value="">-- Pilih Kasir --</option>';
    
    masterKasir.forEach(user => {
        const idKasir = user.kd_ksr || user.kode_member || Object.values(user)[0];
        const namaKasir = user.nama_member || user.nama || idKasir;
        
        if(idKasir) {
            const option = document.createElement("option");
            option.value = JSON.stringify({ kode: idKasir, nama: namaKasir }); // Simpan kode dan nama
            option.textContent = `${idKasir} - ${namaKasir}`;
            select.appendChild(option);
        }
    });
}

// --- Keranjang Belanja ---
function tambahKeKeranjang() {
    const inputKataKunci = document.getElementById("search-product").value.trim().toLowerCase();
    if (!inputKataKunci) return;

    const produk = masterProduk.find(p => 
        String(p.plu).toLowerCase() === inputKataKunci || 
        String(p.barcode).toLowerCase() === inputKataKunci
    );

    if (!produk) {
        alert("Produk tidak ditemukan! Cek PLU/Barcode atau pastikan Master Produk telah di-upload.");
        return;
    }

    const hargaSatuan = parseFloat(produk.price1) || 0;
    const nilaiHemat = parseFloat(produk.disc1) || 0; // Kolom diskon
    // Sesuai request: Ambil deskripsi dari s_descp
    const namaProdukSDescp = produk.s_descp || produk.descp || "Produk Retail"; 
    const kodeBarcode = produk.barcode || "N/A";

    const itemEksis = keranjang.find(item => item.plu === produk.plu);

    if (itemEksis) {
        itemEksis.qty += 1;
        itemEksis.total = itemEksis.qty * (itemEksis.hargaSatuan - itemEksis.nilaiHemat);
    } else {
        keranjang.push({
            plu: produk.plu,
            barcode: kodeBarcode,
            s_descp: namaProdukSDescp,
            hargaSatuan: hargaSatuan,
            qty: 1,
            nilaiHemat: nilaiHemat,
            total: 1 * (hargaSatuan - nilaiHemat)
        });
    }

    document.getElementById("search-product").value = ""; 
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
            <td>${item.s_descp}</td>
            <td>Rp ${item.hargaSatuan.toLocaleString('id-ID')}</td>
            <td>${item.qty}</td>
            <td>Rp ${item.nilaiHemat.toLocaleString('id-ID')}</td>
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
    const uangBayar = parseFloat(document.getElementById("cash-payment").value) || 0;
    const kembalian = uangBayar - grandTotal;
    document.getElementById("cash-return").textContent = kembalian >= 0 ? `Rp ${kembalian.toLocaleString('id-ID')}` : "Rp 0";
}

// --- Checkout & Struk Thermal (Sesuai Syarat Baru) ---
function selesaikanTransaksi() {
    const kasirSelect = document.getElementById("cashier-select");
    if (!kasirSelect.value) { alert("Pilih Kasir terlebih dahulu!"); return; }
    if (keranjang.length === 0) { alert("Keranjang kosong!"); return; }

    const grandTotal = parseFloat(document.getElementById("grand-total").textContent.replace(/[^0-9]/g, '')) || 0;
    const uangBayar = parseFloat(document.getElementById("cash-payment").value) || 0;

    if (uangBayar < grandTotal) { alert("Uang pembayaran kurang!"); return; }

    // Waktu realtime untuk cetak & excel
    const waktuRealtime = new Date();
    const strTglJam = waktuRealtime.toLocaleString('id-ID', { year: 'numeric', month: '2-digit', day: '2-digit', hour: '2-digit', minute: '2-digit', second: '2-digit' });
    
    // Parse data kasir (Kode & Nama)
    const dataKasir = JSON.parse(kasirSelect.value);
    const noReg = document.getElementById("reg-number").value;

    // 1. Simpan ke Laporan Harian (Sheet 1)
    laporanHarian.push({
        "No. Registrasi": noReg,
        "Tanggal & Jam Transaksi": strTglJam, // Realtime di excel
        "Kode Kasir": dataKasir.kode,
        "Nama Kasir": dataKasir.nama,
        "Total Belanja (Rp)": grandTotal,
        "Uang Bayar (Rp)": uangBayar,
        "Kembalian (Rp)": (uangBayar - grandTotal)
    });

    // 2. Simpan Rincian Pembelanjaan (Sheet 2)
    keranjang.forEach(item => {
        laporanDetail.push({
            "No. Registrasi": noReg,
            "Kode Barcode Produk": item.barcode, // Sesuai syarat
            "Nama Produk": item.s_descp,         // Sesuai syarat
            "Qty Belanja": item.qty,             // Sesuai syarat
            "Harga Satuan": item.hargaSatuan,
            "Potongan Promo": item.nilaiHemat,
            "Subtotal": item.total
        });
    });

    // 3. Render HTML untuk Struk Thermal
    buatStrukThermal(noReg, strTglJam, dataKasir.kode, dataKasir.nama, grandTotal, uangBayar);

    // 4. Print otomatis
    window.print();

    // 5. Reset UI Transaksi
    renderLaporanTabel();
    keranjang = [];
    renderKeranjang();
    document.getElementById("cash-payment").value = "";
    nomorRegistrasiHarian++;
    updateNoRegistrasi();
}

function buatStrukThermal(noReg, strTglJam, kodeKasir, namaKasir, total, bayar) {
    const divStruk = document.getElementById("thermal-receipt");
    let htmlBarang = "";
    
    // Perulangan Data Barang untuk Struk (Menampilkan PLU, s_descp, harga, qty, diskon)
    keranjang.forEach(item => {
        htmlBarang += `
            <div style="margin-bottom: 5px;">
                <div style="font-weight: bold;">${item.plu} - ${item.s_descp}</div>
                <div class="receipt-item">
                    <span>${item.qty} x ${item.hargaSatuan.toLocaleString('id-ID')}</span>
                    <span>${(item.qty * item.hargaSatuan).toLocaleString('id-ID')}</span>
                </div>
                ${item.nilaiHemat > 0 ? `
                <div class="receipt-item receipt-text-small">
                    <i>Promo/Hemat:</i>
                    <i>-${(item.nilaiHemat * item.qty).toLocaleString('id-ID')}</i>
                </div>` : ''}
            </div>
        `;
    });

    // Template keseluruhan struk
    divStruk.innerHTML = `
        <div style="text-align: center; margin-bottom: 10px;">
            <h3>TOKO RETAIL</h3>
            <div class="receipt-line"></div>
            <p style="font-size: 11px;">
                No : ${noReg}<br>
                Tgl: ${strTglJam}<br>
                Ksr: ${kodeKasir} / ${namaKasir}
            </p>
        </div>
        <div class="receipt-line"></div>
        
        <div>${htmlBarang}</div>
        
        <div class="receipt-line"></div>
        <div class="receipt-item" style="font-weight: bold; font-size: 14px;">
            <span>TOTAL:</span>
            <span>Rp ${total.toLocaleString('id-ID')}</span>
        </div>
        <div class="receipt-item">
            <span>BAYAR:</span>
            <span>Rp ${bayar.toLocaleString('id-ID')}</span>
        </div>
        <div class="receipt-item">
            <span>KEMBALI:</span>
            <span>Rp ${(bayar - total).toLocaleString('id-ID')}</span>
        </div>
        <div class="receipt-line"></div>
        <div style="text-align: center; margin-top: 10px; font-size: 11px;">
            <p>Terima Kasih<br>Selamat Berbelanja Kembali</p>
        </div>
    `;
}

function renderLaporanTabel() {
    const tbody = document.querySelector("#report-table tbody");
    tbody.innerHTML = "";
    laporanHarian.forEach(trx => {
        const tr = document.createElement("tr");
        tr.innerHTML = `
            <td>${trx["No. Registrasi"]}</td>
            <td>${trx["Tanggal & Jam Transaksi"]}</td>
            <td>${trx["Nama Kasir"]}</td>
            <td>Rp ${trx["Total Belanja (Rp)"].toLocaleString('id-ID')}</td>
        `;
        tbody.appendChild(tr);
    });
}

// --- Ekspor Excel (2 Sheet Sesuai Permintaan) ---
function eksporLaporanKeExcel() {
    if (laporanHarian.length === 0) {
        alert("Belum ada transaksi untuk diekspor!");
        return;
    }

    // Membuat file excel baru
    const workbook = XLSX.utils.book_new();

    // Sheet 1: Laporan Harian (Header Penjualan)
    const worksheet1 = XLSX.utils.json_to_sheet(laporanHarian);
    XLSX.utils.book_append_sheet(workbook, worksheet1, "Laporan_Penjualan");

    // Sheet 2: Rincian Pembelanjaan (Detail Item)
    const worksheet2 = XLSX.utils.json_to_sheet(laporanDetail);
    XLSX.utils.book_append_sheet(workbook, worksheet2, "Rincian_Belanja");

    // Unduh otomatis
    const namaFile = `Rekap_Transaksi_${new Date().toISOString().slice(0,10)}.xlsx`;
    XLSX.writeFile(workbook, namaFile);
}
