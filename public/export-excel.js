document.getElementById('export-excel').addEventListener('click', async function () {
    // Tải template Excel từ server
    try {
        const response = await fetch('./template.xlsx');
        if (!response.ok) throw new Error('Không thể tải template.');
        const buffer = await response.arrayBuffer();

        const workbook = new ExcelJS.Workbook();
        await workbook.xlsx.load(buffer);
        const worksheet = workbook.getWorksheet(1); // Chọn sheet đầu tiên

        if (!orderDetails) {
            alert('Thông tin đơn hàng chưa được tải.');
            return;
        }

        // Điền dữ liệu vào các ô trong Excel
        worksheet.getCell('A3').value = `Số: ${orderDetails.maDonhang || ''}`;
        if (orderDetails.donviPhutrach === "BP. BH1" && orderDetails.phuongThucban !== "Bán chéo") {
            worksheet.getCell('A4').value = 'Kính gửi:';
            worksheet.getCell('C4').value = orderDetails.tenNguoilienhe || '';
            worksheet.getCell('H4').value = 'Ngày phát hành:';
            worksheet.getCell('J4').value = orderDetails.ngayPhatHanh || '';

            worksheet.getCell('A5').value = 'Đơn vị:';
            worksheet.getCell('C5').value = orderDetails.tenTochuc || '';
            worksheet.getCell('H5').value = 'Đơn vị trực thuộc:';
            worksheet.getCell('J5').value = orderDetails.donviPhutrach || '';

            worksheet.getCell('A6').value = 'Địa chỉ:';
            worksheet.getCell('C6').value = orderDetails.diachiChitiet || '';
            worksheet.getCell('H6').value = 'Soạn báo giá:';
            worksheet.getCell('J6').value = orderDetails.tenNhanvien || '';

            worksheet.getCell('A7').value = 'SĐT:';
            worksheet.getCell('C7').value = orderDetails.sdtKhachhang || '';
            worksheet.getCell('H7').value = 'SĐT:';
            worksheet.getCell('J7').value = orderDetails.sdtNhanvien || '';

            worksheet.getCell('A8').value = 'Email:';
            worksheet.getCell('C8').value = orderDetails.emailKhachhang || '';
            worksheet.getCell('H8').value = 'CSKH:';
            worksheet.getCell('J8').value = '1900 0282';

        } else if (orderDetails.donviPhutrach === "BP. BH1" && orderDetails.phuongThucban === "Bán chéo") {
            worksheet.getCell('A4').value = 'Kính gửi:';
            worksheet.getCell('C4').value = orderDetails.tenNguoilienhe || '';
            worksheet.getCell('H4').value = 'Ngày phát hành:';
            worksheet.getCell('J4').value = orderDetails.ngayPhatHanh || '';

            worksheet.getCell('A5').value = 'Đơn vị:';
            worksheet.getCell('C5').value = ''; // Không điền đơn vị
            worksheet.getCell('H5').value = 'Đơn vị trực thuộc:';
            worksheet.getCell('J5').value = orderDetails.donviPhutrach || '';

            worksheet.getCell('A6').value = 'Địa chỉ:';
            worksheet.getCell('C6').value = orderDetails.diachiChitiet || '';
            worksheet.getCell('H6').value = 'Soạn báo giá:';
            worksheet.getCell('J6').value = orderDetails.tenNhanvien || '';

            worksheet.getCell('A7').value = 'SĐT:';
            worksheet.getCell('C7').value = orderDetails.sdtKhachhang || '';
            worksheet.getCell('H7').value = 'SĐT:';
            worksheet.getCell('J7').value = orderDetails.sdtNhanvien || '';

            worksheet.getCell('A8').value = 'Địa chỉ công trình:';
            worksheet.getCell('C8').value = `${orderDetails.tenKhachhangcuoi || ''} - ${orderDetails.diachiKhachhangcuoi || ''} - ${orderDetails.sdtKhachhangcuoi || ''}`;
            worksheet.getCell('H8').value = 'CSKH:';
            worksheet.getCell('J8').value = '1900 0282';

        } else if (orderDetails.donviPhutrach !== "BP. BH1" && orderDetails.hanGiaohang !== "") {
            worksheet.getCell('A4').value = 'Kính gửi:';
            worksheet.getCell('C4').value = ''; // Không điền người nhận
            worksheet.getCell('H4').value = 'Ngày phát hành:';
            worksheet.getCell('J4').value = orderDetails.ngayPhatHanh || '';

            worksheet.getCell('A5').value = 'Đơn vị:';
            worksheet.getCell('C5').value = orderDetails.donviPhutrach || '';
            worksheet.getCell('H5').value = 'Dự kiến giao:';
            worksheet.getCell('J5').value = orderDetails.hanGiaohang || '';

            worksheet.getCell('A6').value = 'Địa chỉ:';
            worksheet.getCell('C6').value = orderDetails.diachi || '';
            worksheet.getCell('H6').value = 'Soạn báo giá:';
            worksheet.getCell('J6').value = orderDetails.tenNhanvien || '';

            worksheet.getCell('A7').value = 'SĐT:';
            worksheet.getCell('C7').value = orderDetails.sdtNhanvien || '';

            worksheet.getCell('A8').value = 'Công trình:';
            worksheet.getCell('C8').value = `${orderDetails.tenNguoilienhe || ''} - ${orderDetails.diachiChitiet || ''} - ${orderDetails.sdtKhachhang || ''}`;
            worksheet.getCell('H8').value = 'CSKH:';
            worksheet.getCell('J8').value = '1900 0282';

        } else if (orderDetails.donviPhutrach !== "BP. BH1" && orderDetails.hanGiaohang === "") {
            worksheet.getCell('A4').value = 'Kính gửi:';
            worksheet.getCell('C4').value = ''; // Không điền người nhận
            worksheet.getCell('H4').value = 'Ngày phát hành:';
            worksheet.getCell('J4').value = orderDetails.ngayPhatHanh || '';

            worksheet.getCell('A5').value = 'Đơn vị:';
            worksheet.getCell('C5').value = orderDetails.donviPhutrach || '';
            worksheet.getCell('H5').value = 'Dự kiến giao:';
            worksheet.getCell('J5').value = 'Trao đổi với QLSX';

            worksheet.getCell('A6').value = 'Địa chỉ:';
            worksheet.getCell('C6').value = orderDetails.diachi || '';
            worksheet.getCell('H6').value = 'Soạn báo giá:';
            worksheet.getCell('J6').value = orderDetails.tenNhanvien || '';

            worksheet.getCell('A7').value = 'SĐT:';
            worksheet.getCell('C7').value = orderDetails.sdtNhanvien || '';

            worksheet.getCell('A8').value = 'Công trình:';
            worksheet.getCell('C8').value = `${orderDetails.tenNguoilienhe || ''} - ${orderDetails.diachiChitiet || ''} - ${orderDetails.sdtKhachhang || ''}`;
            worksheet.getCell('H8').value = 'CSKH:';
            worksheet.getCell('J8').value = '1900 0282';
        }
        worksheet.getCell('H512').value = orderDetails.tongSobo || '';
        worksheet.getCell('L512').value = formatNumber(orderDetails.congnpp || '');
        worksheet.getCell('H513').value = orderDetails.mucChietkhaunpp || '';
        worksheet.getCell('L513').value = formatNumber(orderDetails.giatriChietkhaunpp || '');
        worksheet.getCell('L514').value = formatNumber(orderDetails.phiVanchuyenlapdatnpp || '');
        worksheet.getCell('H515').value = `${orderDetails.mucthueGTGTnpp || ''}%`;
        worksheet.getCell('L515').value = formatNumber(orderDetails.thueGTGTnpp || '');
        worksheet.getCell('L516').value = formatNumber(orderDetails.tamUngnpp || '');
        worksheet.getCell('L517').value = formatNumber(orderDetails.sotienConthieunpp || '');
        worksheet.getCell('A518').value = `Bằng chữ: ${orderDetails.sotienBangchu || ''}`;
        // Điền chi tiết sản phẩm vào Excel
        let startRow = 12; // Ví dụ: bắt đầu từ dòng 12
        orderItems.forEach((item, index) => {
            const row = worksheet.getRow(startRow + index);
            row.getCell(1).value = item.sttTrongdon;
            row.getCell(2).value = item.vitriLapdat;
            row.getCell(3).value = item.maSanphamid;
            row.getCell(4).value = item.diengiai;
            row.getCell(5).value = formatNumber(item.chieuRong);
            row.getCell(6).value = formatNumber(item.chieuCao);
            row.getCell(7).value = item.dienTich;
            row.getCell(8).value = item.soLuong;
            row.getCell(9).value = item.dvt;
            row.getCell(10).value = item.khoiLuong;
            row.getCell(11).value = formatNumber(item.dongianpp);
            row.getCell(12).value = formatNumber(item.giabannpp);
        });
        if (orderDetails.thueGTGTnpp === 0) {
            worksheet.getCell('A519').value = '1. Giá trên đã bao gồm thuế GTGT.';
        } else {
            worksheet.getCell('A519').value = '1. Giá trên chưa bao gồm thuế GTGT.';
        }

        // Điều kiện cho phí vận chuyển và phương thức bán
        if ((orderDetails.phiVanchuyenlapdatnpp === "" || orderDetails.phiVanchuyenlapdatnpp === 0) && orderDetails.phuongThucban !== "Bán lẻ") {
            worksheet.getCell('A520').value = '2. Giá trên chưa bao gồm phí vận chuyển, lắp đặt.';
        } else {
            worksheet.getCell('A520').value = '2. Giá trên đã bao gồm phí vận chuyển, lắp đặt.';
        }

        // Điều kiện cho đơn vị phụ trách và phương thức bán
        if (orderDetails.donviPhutrach === "BP. BH1" && orderDetails.phuongThucban === "Bán đại lý") {
            worksheet.getCell('A521').value = '3. Giá trên có hiệu lực 30 ngày kể từ ngày phát hành.';
            worksheet.getCell('A522').value = '4. Thanh toán 100% tổng giá trị đơn hàng trước khi giao hàng.';
            worksheet.getCell('A523').value = '5. Giao hàng sau 3 đến 5 ngày làm việc, kể từ ngày chốt đơn.';
            worksheet.getCell('A524').value = '6. Thời gian bảo hành 12 tháng theo tiêu chuẩn của nhà sản xuất.';
        } else {
            worksheet.getCell('A521').value = '3. Giá trên có hiệu lực 30 ngày kể từ ngày phát hành.';
            worksheet.getCell('A522').value = '4. Tạm ứng 50% tổng giá trị đơn hàng, thanh toán hết số còn lại sau khi nghiệm thu bàn giao.';
            worksheet.getCell('A523').value = '5. Lắp đặt sau 5 đến 7 ngày làm việc, kể từ ngày nhận được tiền tạm ứng lần 1.';
            worksheet.getCell('A524').value = '6. Thời gian bảo hành:';
            worksheet.getCell('A525').value = ' - Bảo hành 2 năm sản phẩm cửa lưới BS-Polyester, cửa rèm vải visor, cửa rèm tổ ong và phụ kiện lắp đồng bộ với cửa: tay nắm kèm khóa, động cơ cuốn, điều khiển, vật tư nhựa.';
            worksheet.getCell('A526').value = ' - Bảo hành 3 năm sản phẩm cửa lưới sợi thủy tinh, cửa lưới HQ-Polyester, cửa lưới PVC, cửa xếp lưới nhôm, cửa xếp nhựa PC, cửa lưới thép chống cắt.';
            worksheet.getCell('A527').value = ' - Bảo hành 5 năm sản phẩm cửa sử dụng lưới Inox.';
        }

        for (let rowNum = 12; rowNum <= 511; rowNum++) {
            const cellValue = worksheet.getCell(`A${rowNum}`).value;

            // Kiểm tra nếu ô A[rowNum] không có dữ liệu hoặc là trống
            if (cellValue === null || cellValue === '') {
                worksheet.getRow(rowNum).hidden = true; // Ẩn dòng tương ứng
            }
        }

        // Kiểm tra và ẩn các dòng từ L513 đến L516 nếu giá trị trong các ô đó là 0 hoặc trống
        if (worksheet.getCell('L513').value === '0' || worksheet.getCell('L513').value === '') {
            worksheet.getRow(513).hidden = true; // Ẩn dòng 513
        }

        if (worksheet.getCell('L514').value === '0' || worksheet.getCell('L514').value === '') {
            worksheet.getRow(514).hidden = true; // Ẩn dòng 514
        }

        if (worksheet.getCell('L515').value === '0' || worksheet.getCell('L515').value === '') {
            worksheet.getRow(515).hidden = true; // Ẩn dòng 515
        }

        if (worksheet.getCell('L516').value === '0' || worksheet.getCell('L516').value === '') {
            worksheet.getRow(516).hidden = true; // Ẩn dòng 516
        }


        // Lưu file Excel và tải về
        const outputBuffer = await workbook.xlsx.writeBuffer();
        const blob = new Blob([outputBuffer], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
        const link = document.createElement('a');
        link.href = URL.createObjectURL(blob);
        link.download = `Báo giá số ${orderDetails.maDonhang}.xlsx`;

        link.click();
    } catch (error) {
        console.error('Lỗi xuất Excel:', error);
        alert('Không thể xuất Excel. Vui lòng thử lại.');
    }
});