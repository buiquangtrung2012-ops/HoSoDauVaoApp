/**
 * WORD SERVICE (Desktop Version - Pro with Template Filling)
 * Uses PizZip and Docxtemplater for standalone document automation.
 */

export const WordService = {

    isStandalone: () => {
        return typeof Word === 'undefined';
    },

    /**
     * Điền dữ liệu vào file mẫu .docx sử dụng Docxtemplater
     * @param {File} templateFile - File mẫu người dùng chọn
     * @param {Object} state - Dữ liệu hiện tại của App
     */
    fillTemplate: async (templateFile, state) => {
        return new Promise((resolve, reject) => {
            const reader = new FileReader();
            reader.onload = async (e) => {
                try {
                    const content = e.target.result;
                    
                    // Sử dụng require của Node.js (Electron nodeIntegration: true)
                    const PizZip = window.require('pizzip');
                    const Docxtemplater = window.require('docxtemplater');

                    const zip = new PizZip(content);
                    const docSearch = new Docxtemplater(zip, {
                        paragraphLoop: true,
                        linebreaks: true,
                    });

                    // Chuẩn bị dữ liệu để fill
                    // Map dự án
                    const data = { ...state.duAn };
                    
                    // Thêm danh sách Nhân sự
                    data.nhanSu = state.nhanSu.map(row => ({
                        stt: row[0],
                        ten: row[1],
                        chucDanh: row[2],
                        chuyenNganh: row[3],
                        sdt: row[4]
                    }));

                    // Thêm danh sách Máy móc
                    data.mayMoc = state.mayMoc.map(row => ({
                        stt: row[0],
                        ten: row[1],
                        dvt: row[2],
                        sl: row[3],
                        chuSoHuu: row[4],
                        tinhTrang: row[5]
                    }));

                    // Thêm danh sách Vật liệu
                    data.vatLieu = state.vatLieu.map(row => ({
                        stt: row[0],
                        ten: row[1],
                        tieuChuan: row[2],
                        nguonGoc: row[3],
                        ghiChu: row[4]
                    }));

                    // Thêm danh sách Phòng thí nghiệm
                    data.thiNghiem = state.thiNghiem.map(row => ({
                        stt: row[0],
                        dvtn: row[1],
                        diaChi: row[2],
                        tenPtn: row[3],
                        chucNang: row[4]
                    }));

                    // Render (điền dữ liệu)
                    docSearch.render(data);

                    const out = docSearch.getZip().generate({
                        type: "blob",
                        mimeType: "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                    });

                    resolve(out);
                } catch (error) {
                    reject(error);
                }
            };
            reader.onerror = (err) => reject(err);
            reader.readAsArrayBuffer(templateFile);
        });
    },

    // Các hàm mockup cho chế độ standalone để không gây lỗi
    updateDocVariables: async () => { console.log("Standalone: Sync DocVars skip."); },
    updateDocumentVariables: async () => { console.log("Standalone: Sync DocVars skip."); },
    updateAllFields: async () => { console.log("Standalone: Sync Fields skip."); },
    replaceInDocument: async () => { console.log("Standalone: Replace skip."); },
    xuatBang: async () => { console.log("Standalone: Xuat Bang skip."); },
    applyModernStyleToDocument: async () => { console.log("Standalone: Apply Style skip."); }
};
