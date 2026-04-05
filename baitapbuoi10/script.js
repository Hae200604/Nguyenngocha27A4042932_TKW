// ================== CLASS ==================
class SinhVien {
    constructor(maSV, hoTen) {
        this.maSV = maSV;
        this.hoTen = hoTen;
        this.email = taoemail(hoTen, maSV);
        this.khoa = khoa(maSV);
        this.khoaHoc = layKhoaHoc(maSV);
    }
}

// ================== TẠO EMAIL ==================
function taoemail(fullName, studentId){
    let arr = fullName.toLowerCase().split(" ");
    let ten = arr[arr.length - 1]; 
    let ho = arr[0]; 
    let dem= "";

    for(let i=1;i<arr.length-1;i++){
        dem+=arr[i].charAt(0);
    }
    
    let email = ten + ho.charAt(0) + dem + "." + studentId + "@hvnh.edu.vn";
    return email;
}

// ================== KHOA ==================
function khoa(studentId){
    let facultyCode = studentId.substring(2,6);
    let faculty="";
    switch (facultyCode) {
        case "A401":
            faculty = "Ngân hàng"; break;
        case "A402":
            faculty = "Kế toán và kiểm toán"; break;
        case "A403":
            faculty = "Quản trị kinh doanh"; break;
        case "A404":
            faculty = "Công nghệ thông tin và kinh tế số"; break;
        case "A405":
            faculty = "Kinh doanh quốc tế"; break;
        case "A406":
            faculty = "Luật"; break;
        case "A407":
            faculty = "Kinh tế"; break;
        case "A408":
            faculty = "Khoa học dữ liệu"; break;
        case "A751":
            faculty = "Ngôn ngữ"; break;
        default:
            faculty = "Không xác định";
    }
    return faculty;
}

// ================== KHÓA HỌC ==================
function layKhoaHoc(studentId){
    let year = studentId.substring(0,2);
    return "K" + year;
}

// ================== ĐỌC EXCEL ==================
function readExcel(file) {
    const reader = new FileReader();

    reader.onload = function (e) {
        const data = new Uint8Array(e.target.result);
        const workbook = XLSX.read(data, { type: 'array' });

        const sheet = workbook.Sheets[workbook.SheetNames[0]];
        const json = XLSX.utils.sheet_to_json(sheet);

        let ds = [];

        json.forEach(row => {
            let sv = new SinhVien(row["Mã SV"], row["Họ tên"]);
            ds.push(sv);
        });

        hienThi(ds);
    };

    reader.readAsArrayBuffer(file);
}

// ================== HIỂN THỊ ==================
function hienThi(ds) {
    let table = document.getElementById("table");

    let html = `
        <tr>
            <th>Mã SV</th>
            <th>Họ tên</th>
            <th>Email</th>
            <th>Khoa</th>
            <th>Khóa</th>
        </tr>
    `;

    ds.forEach(sv => {
        html += `
            <tr>
                <td>${sv.maSV}</td>
                <td>${sv.hoTen}</td>
                <td>${sv.email}</td>
                <td>${sv.khoa}</td>
                <td>${sv.khoaHoc}</td>
            </tr>
        `;
    });

    table.innerHTML = html;
}