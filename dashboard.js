import React, { useState, useEffect, useCallback, useRef } from "react";
import {
  Chart as ChartJS,
  CategoryScale,
  LinearScale,
  BarElement,
  LineElement,
  PointElement,
  Title,
  Tooltip,
  Legend,
  ArcElement,
} from "chart.js";
import { Bar, Line, Pie } from "react-chartjs-2";
import * as XLSX from "xlsx";
import ExcelJS from "exceljs";
import "./dashboard.css";
import {
  FiBarChart2,
  FiDollarSign,
  FiCreditCard,
  FiPackage,
  FiDownload,
  FiFilter,
  FiCalendar,
  FiTrendingUp,
  FiShoppingCart,
  FiFileText,
  FiRefreshCw,
  FiSettings,
  FiCheckSquare,
  FiSquare,
  FiEye,
  FiPieChart,
  FiSearch,
  FiClipboard,
} from "react-icons/fi";

ChartJS.register(
  CategoryScale,
  LinearScale,
  BarElement,
  LineElement,
  PointElement,
  ArcElement,
  Title,
  Tooltip,
  Legend
);

const Dashboard = () => {
  const [data, setData] = useState({
    THDoanhSo: [],
    THDoanhSoSP: [],
    THHoaDon: [],
    CTBanHang: [],
  });
  const [nhapHangData, setNhapHangData] = useState([]);
  const [khoHangData, setKhoHangData] = useState([]);
  const [shiftData, setShiftData] = useState([]);
  const [kiemKeData, setKiemKeData] = useState([]);
  const [banhKemData, setBanhKemData] = useState([]); // Thêm state cho dữ liệu bánh kem
  const [wholesaleData, setWholesaleData] = useState([]); // Thêm state cho dữ liệu bánh kem đặt
  const [stores, setStores] = useState([]);
  // Loading states riêng cho từng loại dữ liệu
  const [loadingState, setLoadingState] = useState({
    thongKe: false,
    nhapHang: false,
    ketCa: false,
    banhKem: false,
    wholesale: false,
    kiemKe: false,
    khoHang: false,
  });
  const [loadingStores, setLoadingStores] = useState(true);
  const today = new Date().toISOString().split("T")[0];

  // Helper: cập nhật loading state cho 1 key
  const setLoadingFor = (key, value) => {
    setLoadingState((prev) => ({ ...prev, [key]: value }));
  };
  // Computed: có đang loading bất kỳ API nào không
  const isAnyLoading = Object.values(loadingState).some(Boolean);

  const [filters, setFilters] = useState({
    tuNgay: today,
    denNgay: today,
    dsMaCH: [],
  });

  const [exportOptions, setExportOptions] = useState({
    doanhSo: true,
    doanhSoSP: true,
    hoaDon: true,
    chiTietBanHang: true,
    nhapHang: true,
    tonKho: true,
    ketCa: true,
    kiemKe: true,
    banhKem: true, // Thêm option cho bánh kem
    wholesale: true, // Thêm option cho bánh kem đặt
  });

  const [showExportDropdown, setShowExportDropdown] = useState(false);
  const [storeSearch, setStoreSearch] = useState("");
  const [showSplitXuat, setShowSplitXuat] = useState(false);
  const [showSplitGiaTriXuat, setShowSplitGiaTriXuat] = useState(false);

  // Fetch stores from API
  const fetchStores = async () => {
    try {
      const response = await fetch("https://pos.doanquochoa.name.vn/stores");
      const storesData = await response.json();

      // Check admin status and user's store ID from localStorage
      const isBakery = localStorage.getItem("isBakery") === "true";
      const isAdminHCM = localStorage.getItem("HCM_ADMIN") === "true"
      const userStoreIdStr = localStorage.getItem("IDCH");
      const userStoreId = userStoreIdStr ? parseInt(userStoreIdStr, 10) : null;

      let filteredStores = storesData;

      // If not admin, filter to only show user's store
      if (isBakery && userStoreId) {
        filteredStores = storesData.filter((store) => store.IDCH === userStoreId);
      } else if (isAdminHCM) {
        filteredStores = storesData.filter((store) =>  store.LoaiCH === 3 || store.LoaiCH === "3");
      }

      setStores(filteredStores);
      // Initialize with first few stores selected
      /*
      if (storesData.length > 0) {
        const defaultStores = storesData.slice(0, 3).map((store) => store.IDCH);
        setFilters((prev) => ({
          ...prev,
          dsMaCH: defaultStores,
        }));
      }
      */
    } catch (error) {
      console.error("Lỗi tải danh sách cửa hàng:", error);
    } finally {
      setLoadingStores(false);
    }
  };

  const fetchData = () => {
    if (filters.dsMaCH.length === 0) {
      alert("Vui lòng chọn ít nhất một cửa hàng");
      return;
    }

    // Reset kho hàng data khi fetch lại (sẽ chỉ load khi xuất báo cáo)
    setKhoHangData([]);

    // 1) Thống kê (doanh số, doanh số SP, hóa đơn, chi tiết bán hàng)
    setLoadingFor("thongKe", true);
    fetch("https://pos.doanquochoa.name.vn/api/thong-ke", {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify(filters),
    })
      .then((res) => res.json())
      .then((result) => setData(result))
      .catch((err) => console.error("Lỗi tải thống kê:", err))
      .finally(() => setLoadingFor("thongKe", false));

    // 2) Nhập hàng
    setLoadingFor("nhapHang", true);
    const nhapHangParams = new URLSearchParams({
      MaCHs: filters.dsMaCH.join(","),
      startDate: filters.tuNgay,
      endDate: filters.denNgay,
    });
    fetch(`https://pos.doanquochoa.name.vn/api/sale-nhap-hang-by-ch?${nhapHangParams}`)
      .then((res) => res.json())
      .then((nhapHangResult) => {
        const sorted = nhapHangResult.sort((a, b) => (parseInt(a.SoPN) || 0) - (parseInt(b.SoPN) || 0));
        setNhapHangData(sorted);
      })
      .catch((err) => console.error("Lỗi tải nhập hàng:", err))
      .finally(() => setLoadingFor("nhapHang", false));

    // 3) Kết ca
    setLoadingFor("ketCa", true);
    const shiftParams = new URLSearchParams();
    if (filters.tuNgay) shiftParams.append("fromDate", filters.tuNgay);
    if (filters.denNgay) shiftParams.append("toDate", filters.denNgay);
    if (filters.dsMaCH && filters.dsMaCH.length > 0) {
      shiftParams.append("idch", filters.dsMaCH.join(","));
    }
    fetch(`https://pos.doanquochoa.name.vn/api/sale-ket-ca?${shiftParams.toString()}`)
      .then((res) => res.json())
      .then((result) => setShiftData(Array.isArray(result) ? result : []))
      .catch((err) => console.error("Lỗi tải kết ca:", err))
      .finally(() => setLoadingFor("ketCa", false));

    // 4) Bánh kem bán - gọi tất cả cửa hàng song song
    setLoadingFor("banhKem", true);
    Promise.all(
      filters.dsMaCH.map((storeId) =>
        fetch("https://pos.doanquochoa.name.vn/api/doansobanhkem", {
          method: "POST",
          headers: { "Content-Type": "application/json" },
          body: JSON.stringify({
            MaCH: storeId,
            NgayBatDau: `${filters.tuNgay}T00:00:00`,
            NgayKetThuc: `${filters.denNgay}T23:59:59`,
          }),
        })
          .then((res) => (res.ok ? res.json() : null))
          .then((r) => (r && r.success && r.data ? r.data : []))
          .catch(() => [])
      )
    )
      .then((results) => setBanhKemData(results.flat()))
      .finally(() => setLoadingFor("banhKem", false));

    // 5) Bánh kem đặt - gọi tất cả cửa hàng song song
    setLoadingFor("wholesale", true);
    Promise.all(
      filters.dsMaCH.map((storeId) =>
        fetch("https://pos.doanquochoa.name.vn/api/cake-wholesale-summary-detail", {
          method: "POST",
          headers: { "Content-Type": "application/json" },
          body: JSON.stringify({
            MaCH: storeId,
            NgayBatDau: filters.tuNgay,
            NgayKetThuc: filters.denNgay,
          }),
        })
          .then((res) => (res.ok ? res.json() : null))
          .then((r) => (Array.isArray(r) ? r : []))
          .catch(() => [])
      )
    )
      .then((results) => setWholesaleData(results.flat()))
      .finally(() => setLoadingFor("wholesale", false));

    // 6) Phiếu kiểm kê
    setLoadingFor("kiemKe", true);
    fetch("https://pos.doanquochoa.name.vn/api/phieu-kiem-ke")
      .then((res) => (res.ok ? res.json() : []))
      .then((kiemKeResult) => {
        const filtered = kiemKeResult.filter((phieu) => {
          const phieuDate = new Date(phieu.NgayKiemKe);
          const fromDate = new Date(filters.tuNgay);
          const toDate = new Date(filters.denNgay);
          return phieuDate >= fromDate && phieuDate <= toDate;
        });
        setKiemKeData(filtered);
      })
      .catch((err) => {
        console.error("Lỗi tải kiểm kê:", err);
        setKiemKeData([]);
      })
      .finally(() => setLoadingFor("kiemKe", false));
  };

  useEffect(() => {
    fetchStores();
  }, []);

  // Chỉ gọi fetchData một lần khi stores đã load xong và có cửa hàng được chọn
  const hasAutoFetched = useRef(false);
  useEffect(() => {
    if (!loadingStores && filters.dsMaCH.length > 0 && !hasAutoFetched.current) {
      hasAutoFetched.current = true;
      fetchData();
    }
  }, [loadingStores, filters.dsMaCH]);

  const handleFilterChange = (field, value) => {
    setFilters((prev) => ({
      ...prev,
      [field]: value,
    }));
  };

  const handleStoreChange = (storeId, checked) => {
    setFilters((prev) => ({
      ...prev,
      dsMaCH: checked
        ? [...prev.dsMaCH, storeId]
        : prev.dsMaCH.filter((id) => id !== storeId),
    }));
  };

  const handleSelectAllStores = (checked) => {
    setFilters((prev) => ({
      ...prev,
      dsMaCH: checked ? stores.map((store) => store.IDCH) : [],
    }));
  };

  const handleExportOptionChange = (option, checked) => {
    setExportOptions((prev) => ({
      ...prev,
      [option]: checked,
    }));
  };

  const handleSelectAllExports = (checked) => {
    setExportOptions({
      doanhSo: checked,
      doanhSoSP: checked,
      hoaDon: checked,
      chiTietBanHang: checked,
      nhapHang: checked,
      tonKho: checked,
      ketCa: checked,
      kiemKe: checked,
      banhKem: checked, // Thêm bánh kem
      wholesale: checked, // Thêm bánh kem đặt
    });
  };

  const toggleExportDropdown = () => {
    setShowExportDropdown(!showExportDropdown);
  };

  // Xuất Doanh Số
  const exportDoanhSo = async () => {
    const workbook = new ExcelJS.Workbook();
    const worksheet = workbook.addWorksheet("Doanh Số");

    // Thêm tiêu đề lớn
    worksheet.mergeCells("A1:E1");
    worksheet.getCell("A1").value = "BÁO CÁO DOANH SỐ";
    worksheet.getCell("A1").font = { size: 16, bold: true };
    worksheet.getCell("A1").alignment = {
      vertical: "middle",
      horizontal: "center",
    };

    worksheet.mergeCells("A2:E2");
    worksheet.getCell(
      "A2"
    ).value = `Từ ngày ${filters.tuNgay} đến ngày ${filters.denNgay}`;
    worksheet.getCell("A2").font = { italic: true };
    worksheet.getCell("A2").alignment = {
      vertical: "middle",
      horizontal: "center",
    };

    worksheet.addRow([]);

    // Thêm tiêu đề cột
    const headerRow = worksheet.addRow([
      "Ngày",
      "Cửa Hàng",
      "Doanh Thu",
      "Giảm Giá",
      "Còn Lại",
    ]);
    headerRow.font = { bold: true };
    headerRow.alignment = { horizontal: "center" };
    headerRow.fill = {
      type: "pattern",
      pattern: "solid",
      fgColor: { argb: "FFDDEEFF" },
    };

    // Set column widths
    worksheet.getColumn(1).width = 15;
    worksheet.getColumn(2).width = 30;
    worksheet.getColumn(3).width = 18;
    worksheet.getColumn(4).width = 15;
    worksheet.getColumn(5).width = 18;

    let totalDoanhThu = 0;
    let totalGiamGia = 0;
    let totalConLai = 0;

    data.THDoanhSo.forEach((item) => {
      worksheet.addRow([
        new Date(item.NgayThangNam).toLocaleDateString("vi-VN"),
        item.TenCuaHang,
        item.DoanhThu,
        item.GiamGia,
        item.DoanhThuConLai,
      ]);
      totalDoanhThu += item.DoanhThu;
      totalGiamGia += item.GiamGia;
      totalConLai += item.DoanhThuConLai;
    });

    // Thêm hàng tổng cộng
    const totalRow = worksheet.addRow([
      "",
      "TỔNG CỘNG",
      totalDoanhThu,
      totalGiamGia,
      totalConLai,
    ]);
    totalRow.font = { bold: true };
    totalRow.fill = {
      type: "pattern",
      pattern: "solid",
      fgColor: { argb: "FFFFCCCC" },
    };

    // Format numbers
    worksheet.eachRow((row, rowNumber) => {
      if (rowNumber > 4) {
        row.getCell(3).numFmt = "#,##0";
        row.getCell(4).numFmt = "#,##0";
        row.getCell(5).numFmt = "#,##0";
      }
    });

    const buffer = await workbook.xlsx.writeBuffer();
    const blob = new Blob([buffer], {
      type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    });
    const url = window.URL.createObjectURL(blob);
    const a = document.createElement("a");
    a.href = url;
    a.download = `DoanhSo_${filters.tuNgay}_${filters.denNgay}.xlsx`;
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
    window.URL.revokeObjectURL(url);
  };

  // Xuất Doanh Số Sản Phẩm
  const exportDoanhSoSP = async () => {
    const workbook = new ExcelJS.Workbook();
    const worksheet = workbook.addWorksheet("Doanh Số SP");

    // Thêm tiêu đề lớn
    worksheet.mergeCells("A1:E1");
    worksheet.getCell("A1").value = "BÁO CÁO DOANH SỐ SẢN PHẨM";
    worksheet.getCell("A1").font = { size: 16, bold: true };
    worksheet.getCell("A1").alignment = {
      vertical: "middle",
      horizontal: "center",
    };

    worksheet.mergeCells("A2:F2");
    worksheet.getCell(
      "A2"
    ).value = `Từ ngày ${filters.tuNgay} đến ngày ${filters.denNgay}`;
    worksheet.getCell("A2").font = { italic: true };
    worksheet.getCell("A2").alignment = {
      vertical: "middle",
      horizontal: "center",
    };

    worksheet.addRow([]);

    // Thêm tiêu đề cột
    const headerRow = worksheet.addRow([
      "Mã SP",
      "MaNLSP",
      "Tên Sản Phẩm",
      "Số Lượng",
      "Giá Bán",
      "Thành Tiền",
    ]);
    headerRow.font = { bold: true };
    headerRow.alignment = { horizontal: "center" };
    headerRow.fill = {
      type: "pattern",
      pattern: "solid",
      fgColor: { argb: "FFDDEEFF" },
    };

    // Set column widths
    worksheet.getColumn(1).width = 15;
    worksheet.getColumn(2).width = 15;
    worksheet.getColumn(3).width = 40;
    worksheet.getColumn(4).width = 15;
    worksheet.getColumn(5).width = 15;
    worksheet.getColumn(6).width = 18;

    let totalSoLuong = 0;
    let totalThanhTien = 0;

    data.THDoanhSoSP.forEach((item) => {
      worksheet.addRow([
        item.MaSP,
        item.MaNLSP,
        item.TenSP,
        item.SoLuong,
        item.GiaBan,
        item.ThanhTien,
      ]);
      totalSoLuong += item.SoLuong;
      totalThanhTien += item.ThanhTien;
    });

    // Thêm hàng tổng cộng
    const totalRow = worksheet.addRow([
      "",
      "",
      "TỔNG CỘNG",
      totalSoLuong,
      "",
      totalThanhTien,
    ]);
    totalRow.font = { bold: true };
    totalRow.fill = {
      type: "pattern",
      pattern: "solid",
      fgColor: { argb: "FFFFCCCC" },
    };

    // Format numbers
    worksheet.eachRow((row, rowNumber) => {
      if (rowNumber > 4) {
        row.getCell(4).numFmt = "#,##0";
        row.getCell(5).numFmt = "#,##0";
        row.getCell(6).numFmt = "#,##0";
      }
    });

    const buffer = await workbook.xlsx.writeBuffer();
    const blob = new Blob([buffer], {
      type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    });
    const url = window.URL.createObjectURL(blob);
    const a = document.createElement("a");
    a.href = url;
    a.download = `DoanhSoSanPham_${filters.tuNgay}_${filters.denNgay}.xlsx`;
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
    window.URL.revokeObjectURL(url);
  };

  // Xuất Hóa Đơn
  const exportHoaDon = async () => {
    const workbook = new ExcelJS.Workbook();
    const worksheet = workbook.addWorksheet("Hóa Đơn");

    if (data.THHoaDon.length === 0) {
      worksheet.addRow(["Không có dữ liệu hóa đơn"]);
    } else {
      const headerKeys = Object.keys(data.THHoaDon[0]);
      const headerCount = headerKeys.length;

      // Thêm tiêu đề lớn
      worksheet.mergeCells(1, 1, 1, headerCount);
      worksheet.getCell(1, 1).value = "BÁO CÁO HÓA ĐƠN";
      worksheet.getCell(1, 1).font = { size: 16, bold: true };
      worksheet.getCell(1, 1).alignment = {
        vertical: "middle",
        horizontal: "center",
      };

      worksheet.mergeCells(2, 1, 2, headerCount);
      worksheet.getCell(
        2,
        1
      ).value = `Từ ngày ${filters.tuNgay} đến ngày ${filters.denNgay}`;
      worksheet.getCell(2, 1).font = { italic: true };
      worksheet.getCell(2, 1).alignment = {
        vertical: "middle",
        horizontal: "center",
      };

      worksheet.addRow([]);

      // Thêm tiêu đề cột
      const headerRow = worksheet.addRow(headerKeys);
      headerRow.font = { bold: true };
      headerRow.alignment = { horizontal: "center" };
      headerRow.fill = {
        type: "pattern",
        pattern: "solid",
        fgColor: { argb: "FFDDEEFF" },
      };

      // Set column widths
      headerKeys.forEach((key, index) => {
        worksheet.getColumn(index + 1).width = 20;
      });

      // Calculate totals for monetary columns
      let totals = {};
      const monetaryColumns = [
        "DoanhThu",
        "ThanhTien",
        "TongTien",
        "GiamGia",
        "TienThue",
        "ThanhTienConLai",
      ];

      data.THHoaDon.forEach((item) => {
        const rowData = headerKeys.map((key) => item[key]);
        worksheet.addRow(rowData);

        // Calculate totals for monetary columns
        headerKeys.forEach((key) => {
          if (monetaryColumns.includes(key) && typeof item[key] === "number") {
            totals[key] = (totals[key] || 0) + item[key];
          }
        });
      });

      // Thêm hàng tổng cộng
      const totalRowData = headerKeys.map((key, index) => {
        if (index === 0) return "TỔNG CỘNG";
        if (index === 1) return `${data.THHoaDon.length} hóa đơn`;
        if (totals[key]) return totals[key];
        return "";
      });

      const totalRow = worksheet.addRow(totalRowData);
      totalRow.font = { bold: true };
      totalRow.fill = {
        type: "pattern",
        pattern: "solid",
        fgColor: { argb: "FFFFCCCC" },
      };

      // Format monetary columns
      worksheet.eachRow((row, rowNumber) => {
        if (rowNumber > 4) {
          headerKeys.forEach((key, index) => {
            if (monetaryColumns.includes(key)) {
              row.getCell(index + 1).numFmt = "#,##0";
            }
          });
        }
      });
    }

    const buffer = await workbook.xlsx.writeBuffer();
    const blob = new Blob([buffer], {
      type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    });
    const url = window.URL.createObjectURL(blob);
    const a = document.createElement("a");
    a.href = url;
    a.download = `HoaDon_${filters.tuNgay}_${filters.denNgay}.xlsx`;
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
    window.URL.revokeObjectURL(url);
  };

  // Xuất Chi Tiết Bán Hàng
  const exportChiTietBanHang = async () => {
    const workbook = new ExcelJS.Workbook();
    const worksheet = workbook.addWorksheet("Chi Tiết Bán Hàng");

    if (data.CTBanHang.length === 0) {
      worksheet.addRow(["Không có dữ liệu chi tiết bán hàng"]);
    } else {
      const invoiceMap = {};
      if (data.THHoaDon && data.THHoaDon.length > 0) {
        data.THHoaDon.forEach((hd) => {
          const keys = [
            hd.IDPhieu,
            hd.MaHD,
            hd.SoHD,
            hd.SoHoaDon,
          ].filter(Boolean);
          const value =
            hd.MaHTTT || hd.HinhThucThanhToan || hd.PhuongThuc || "";
          keys.forEach((k) => {
            invoiceMap[k] = value;
          });
        });
      }

      const rawHeaderKeys = Object.keys(data.CTBanHang[0]);
      const headerKeys = rawHeaderKeys.includes("MaHTTT")
        ? rawHeaderKeys
        : [...rawHeaderKeys, "MaHTTT"];

      const headerCount = headerKeys.length;

      // worksheet.mergeCells(1, 1, 1, headerCount);

      // Thêm tiêu đề lớn
      worksheet.mergeCells(1, 1, 1, headerCount);
      worksheet.getCell(1, 1).value = "BÁO CÁO CHI TIẾT BÁN HÀNG";
      worksheet.getCell(1, 1).font = { size: 16, bold: true };
      worksheet.getCell(1, 1).alignment = {
        vertical: "middle",
        horizontal: "center",
      };

      worksheet.mergeCells(2, 1, 2, headerCount);
      worksheet.getCell(
        2,
        1
      ).value = `Từ ngày ${filters.tuNgay} đến ngày ${filters.denNgay}`;
      worksheet.getCell(2, 1).font = { italic: true };
      worksheet.getCell(2, 1).alignment = {
        vertical: "middle",
        horizontal: "center",
      };

      worksheet.addRow([]);

      // Thêm tiêu đề cột
      const headerRow = worksheet.addRow(headerKeys);
      headerRow.font = { bold: true };
      headerRow.alignment = { horizontal: "center" };
      headerRow.fill = {
        type: "pattern",
        pattern: "solid",
        fgColor: { argb: "FFDDEEFF" },
      };

      // Set column widths
      headerKeys.forEach((key, index) => {
        worksheet.getColumn(index + 1).width = 20;
      });

      // Calculate totals for monetary columns
      let totals = {};
      const monetaryColumns = [
        "DoanhThu",
        "ThanhTien",
        "TongTien",
        "TienThue",
        "GiaBan",
        "ThanhTienConLai",
      ];
      const quantityColumns = ["SoLuong"];
      let totalGiamGiaAmount = 0; // Separate calculation for discount amount

      data.CTBanHang.forEach((item) => {
        // Fix ThanhTien calculation: ThanhTien = SoLuong * GiaBan
        const correctedItem = { ...item };
        if (correctedItem.SoLuong && correctedItem.GiaBan) {
          correctedItem.ThanhTien =
            correctedItem.SoLuong * correctedItem.GiaBan;
        }
        if (!correctedItem.MaHTTT) {
          const keyCandidates = [
            correctedItem.MaHoaDon,
            correctedItem.IDPhieu,
            correctedItem.SoHD,
            correctedItem.SoHoaDon,
          ].filter(Boolean);
          const foundKey = keyCandidates.find((k) => invoiceMap[k] !== undefined);
          if (foundKey) {
            correctedItem.MaHTTT = invoiceMap[foundKey];
          }
        }

        const rowData = headerKeys.map((key) => correctedItem[key]);
        worksheet.addRow(rowData);

        // Calculate totals using corrected values
        headerKeys.forEach((key) => {
          if (
            monetaryColumns.includes(key) &&
            typeof correctedItem[key] === "number"
          ) {
            totals[key] = (totals[key] || 0) + correctedItem[key];
          } else if (
            quantityColumns.includes(key) &&
            typeof correctedItem[key] === "number"
          ) {
            totals[key] = (totals[key] || 0) + correctedItem[key];
          }
        });

        // Calculate discount amount: GiamGia% * ThanhTien
        if (correctedItem.GiamGia && correctedItem.ThanhTien) {
          totalGiamGiaAmount +=
            (correctedItem.GiamGia / 100) * correctedItem.ThanhTien;
        }
      });

      // Thêm hàng tổng cộng
      const totalRowData = headerKeys.map((key, index) => {
        if (index === 0) return "TỔNG CỘNG";
        if (index === 1) return `${data.CTBanHang.length} giao dịch`;
        if (key === "GiamGia") return totalGiamGiaAmount; // Use calculated discount amount
        if (totals[key]) return totals[key];
        return "";
      });

      const totalRow = worksheet.addRow(totalRowData);
      totalRow.font = { bold: true };
      totalRow.fill = {
        type: "pattern",
        pattern: "solid",
        fgColor: { argb: "FFFFCCCC" },
      };

      // Format monetary and quantity columns
      worksheet.eachRow((row, rowNumber) => {
        if (rowNumber > 4) {
          headerKeys.forEach((key, index) => {
            if (
              monetaryColumns.includes(key) ||
              quantityColumns.includes(key)
            ) {
              row.getCell(index + 1).numFmt = "#,##0";
            }
          });
        }
      });
    }

    const buffer = await workbook.xlsx.writeBuffer();
    const blob = new Blob([buffer], {
      type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    });
    const url = window.URL.createObjectURL(blob);
    const a = document.createElement("a");
    a.href = url;
    a.download = `ChiTietBanHang_${filters.tuNgay}_${filters.denNgay}.xlsx`;
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
    window.URL.revokeObjectURL(url);
  };

  const exportNhapHang = async () => {
    // Sử dụng ExcelJS thay vì XLSX
    const workbook = new ExcelJS.Workbook();
    const worksheet = workbook.addWorksheet("Sổ chi tiết nhập hàng");

    // Định nghĩa cột
    worksheet.columns = [
      { header: "Mã SP", key: "MaSP", width: 15 },
      { header: "Tên SP", key: "TenSP", width: 40 },
      { header: "Số lượng", key: "SoLuong", width: 12 },
      { header: "Đơn giá", key: "GiaMua", width: 15 },
      { header: "Giảm giá(%)", key: "GiamPT", width: 15 },
      { header: "Thành tiền", key: "ThanhTien", width: 15 },
      { header: "Ghi chú", key: "GhiChu", width: 20 },
    ];

    // Thêm tiêu đề lớn
    worksheet.mergeCells("A1:G1");
    worksheet.getCell("A1").value = "SỔ CHI TIẾT NHẬP HÀNG";
    worksheet.getCell("A1").font = { size: 16, bold: true };
    worksheet.getCell("A1").alignment = {
      vertical: "middle",
      horizontal: "center",
    };

    worksheet.mergeCells("A2:G2");
    worksheet.getCell(
      "A2"
    ).value = `Từ ngày ${filters.tuNgay} đến ngày ${filters.denNgay}`;
    worksheet.getCell("A2").font = { italic: true };
    worksheet.getCell("A2").alignment = {
      vertical: "middle",
      horizontal: "center",
    };

    worksheet.addRow([]);

    // Gom nhóm dữ liệu theo cửa hàng và ngày
    const consolidatedData = {};
    let totalSoLuong = 0;
    let totalThanhTien = 0;

    nhapHangData.forEach((phieu) => {
      const dateKey = new Date(phieu.NgayNhap).toLocaleDateString("vi-VN");
      const storeKey = phieu.MaCH;
      const key = `${storeKey}_${dateKey}`;
      if (!consolidatedData[key]) {
        consolidatedData[key] = {
          MaCH: storeKey,
          NgayNhap: dateKey,
          Items: {},
        };
      }
      phieu.Items.forEach((item) => {
        const itemKey = item.MaSP;
        if (consolidatedData[key].Items[itemKey]) {
          consolidatedData[key].Items[itemKey].SoLuong += item.SoLuong;
          consolidatedData[key].Items[itemKey].ThanhTien += item.ThanhTien;
        } else {
          consolidatedData[key].Items[itemKey] = { ...item };
        }
        totalSoLuong += item.SoLuong;
        totalThanhTien += item.ThanhTien;
      });
    });

    // Sắp xếp nhóm
    const sortedConsolidatedData = Object.values(consolidatedData).sort(
      (a, b) => {
        if (a.MaCH !== b.MaCH) return a.MaCH.localeCompare(b.MaCH);
        return (
          new Date(a.NgayNhap.split("/").reverse().join("-")) -
          new Date(b.NgayNhap.split("/").reverse().join("-"))
        );
      }
    );

    // Thêm dữ liệu từng nhóm
    let rowIdx = worksheet.lastRow.number + 1;
    sortedConsolidatedData.forEach((group) => {
      // Header nhóm
      worksheet.addRow([
        `Ngày: ${group.NgayNhap}`,
        `Cửa hàng: ${group.MaCH}`,
        "",
        "",
        "",
        "",
        "",
      ]);
      worksheet.getRow(rowIdx).font = { bold: true };
      worksheet.getRow(rowIdx).fill = {
        type: "pattern",
        pattern: "solid",
        fgColor: { argb: "FFEFEFEF" },
      };
      rowIdx++;

      // Sắp xếp sản phẩm
      const sortedItems = Object.values(group.Items).sort((a, b) =>
        a.MaSP.localeCompare(b.MaSP)
      );
      sortedItems.forEach((item) => {
        worksheet.addRow({
          MaSP: item.MaSP,
          TenSP: item.TenSP,
          SoLuong: item.SoLuong,
          GiaMua: item.GiaMua,
          GiamPT: item.GiamPT || 0,
          ThanhTien: item.ThanhTien,
          GhiChu: "",
        });
        rowIdx++;
      });

      worksheet.addRow([]);
      rowIdx++;
    });

    // Thêm hàng tổng cộng
    worksheet.addRow([]);
    const totalRow = worksheet.addRow({
      MaSP: "",
      TenSP: "TỔNG CỘNG",
      SoLuong: totalSoLuong,
      GiaMua: "",
      GiamPT: "",
      ThanhTien: totalThanhTien,
      GhiChu: "",
    });
    totalRow.font = { bold: true };
    totalRow.fill = {
      type: "pattern",
      pattern: "solid",
      fgColor: { argb: "FFFFCCCC" },
    };

    // Định dạng số và tiền tệ
    worksheet.eachRow((row, rowNumber) => {
      if (rowNumber > 4) {
        row.getCell("SoLuong").numFmt = "#,##0";
        row.getCell("GiaMua").numFmt = "#,##0";
        row.getCell("ThanhTien").numFmt = "#,##0";
        row.getCell("GiamPT").numFmt = "0%";
      }
    });

    // Xuất file
    const buffer = await workbook.xlsx.writeBuffer();
    const blob = new Blob([buffer], {
      type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    });
    const url = window.URL.createObjectURL(blob);
    const a = document.createElement("a");
    a.href = url;
    a.download = `SoChiTietNhapHang_${filters.tuNgay}_${filters.denNgay}.xlsx`;
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
    window.URL.revokeObjectURL(url);
  };

  // Hàm fetch dữ liệu kho hàng riêng biệt - chỉ gọi khi cần xuất báo cáo
  const fetchKhoHang = async () => {
    if (filters.dsMaCH.length === 0) {
      alert("Vui lòng chọn ít nhất một cửa hàng");
      return [];
    }
    setLoadingFor("khoHang", true);
    try {
      const tonKhoParams = new URLSearchParams({
        MaCH: filters.dsMaCH[0],
        NgayBatDau: filters.tuNgay,
        NgayKetThuc: filters.denNgay,
      });
      const tonKhoResponse = await fetch(
        `https://pos.doanquochoa.name.vn/api/khohang?${tonKhoParams}`
      );
      const tonKhoResult = await tonKhoResponse.json();
      if (tonKhoResult.success) {
        setKhoHangData(tonKhoResult.data || []);
        return tonKhoResult.data || [];
      } else {
        setKhoHangData([]);
        return [];
      }
    } catch (error) {
      console.error("Lỗi tải dữ liệu tồn kho:", error);
      setKhoHangData([]);
      return [];
    } finally {
      setLoadingFor("khoHang", false);
    }
  };

  // Xuất Tồn Kho
  const exportTonKho = async () => {
    // Fetch dữ liệu kho hàng trước khi xuất
    let exportData = khoHangData;
    if (khoHangData.length === 0) {
      exportData = await fetchKhoHang();
    }
    if (!exportData || exportData.length === 0) {
      alert("Không có dữ liệu tồn kho để xuất báo cáo");
      return;
    }

    const workbook = new ExcelJS.Workbook();
    const worksheet = workbook.addWorksheet("Tồn kho");

    // Thiết lập chiều rộng cột, bổ sung các cột điều phối/điều chuyển
    worksheet.columns = [
      { width: 8 }, // STT
      { width: 15 }, // Mã sản phẩm
      { width: 35 }, // Tên sản phẩm
      { width: 10 }, // ĐVT
      { width: 15 }, // Đơn giá
      { width: 12 }, // Tồn đầu kỳ
      { width: 15 }, // GT đầu kỳ
      { width: 12 }, // Tổng nhập
      { width: 15 }, // GT nhập
      { width: 12 }, // SL điều phối
      { width: 15 }, // GT điều phối
      { width: 12 }, // SL điều chuyển
      { width: 15 }, // GT điều chuyển
      { width: 12 }, // Tổng xuất (bán)
      { width: 15 }, // GT xuất (bán)
      { width: 12 }, // Tổng hủy
      { width: 15 }, // GT hủy
      { width: 12 }, // Tồn cuối kỳ
      { width: 15 }, // GT cuối kỳ
    ];

    // Tiêu đề chính
    worksheet.mergeCells("A1:S1");
    const titleCell = worksheet.getCell("A1");
    titleCell.value = "PHIẾU TỔNG HỢP TỒN KHO";
    titleCell.font = { size: 16, bold: true };
    titleCell.alignment = { horizontal: "center", vertical: "middle" };

    // Thời gian báo cáo
    worksheet.mergeCells("A2:S2");
    const dateCell = worksheet.getCell("A2");
    dateCell.value = `Từ ngày ${filters.tuNgay} đến ngày ${filters.denNgay}`;
    dateCell.font = { size: 12, italic: true };
    dateCell.alignment = { horizontal: "center", vertical: "middle" };

    worksheet.addRow([]);

    // Tiêu đề cột
    const headerRow = worksheet.addRow([
      "STT",
      "Mã sản phẩm",
      "Tên sản phẩm",
      "ĐVT",
      "Đơn giá",
      "Tồn đầu kỳ",
      "GT đầu kỳ",
      "Tổng nhập",
      "GT nhập",
      "SL điều phối",
      "GT điều phối",
      "SL điều chuyển",
      "GT điều chuyển",
      "Tổng xuất (bán)",
      "GT xuất (bán)",
      "Tổng hủy",
      "GT hủy",
      "Tồn cuối kỳ",
      "GT cuối kỳ",
    ]);
    headerRow.font = { bold: true };
    headerRow.alignment = { horizontal: "center", vertical: "middle" };
    headerRow.fill = {
      type: "pattern",
      pattern: "solid",
      fgColor: { argb: "FFE6E6FA" },
    };

    // Border cho tiêu đề
    for (let col = 1; col <= 19; col++) {
      headerRow.getCell(col).border = {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      };
    }

    // Thêm dữ liệu
    let totalTonDK = 0,
      totalGiaTriDK = 0;
    let totalTongNhap = 0,
      totalGiaTriNhap = 0;
    let totalDieuPhoi = 0,
      totalGiaTriDieuPhoi = 0;
    let totalDieuChuyen = 0,
      totalGiaTriDieuChuyen = 0;
    let totalTongXuat = 0,
      totalGiaTriXuat = 0;
    let totalTongHuy = 0,
      totalGiaTriHuy = 0;
    let totalTonCK = 0,
      totalGiaTriCK = 0;

    exportData.forEach((item, index) => {
      // Lấy thông tin điều phối/điều chuyển nếu có
      const dieuPhoi = item.SoLuongDieuPhoi || 0;
      const giaTriDieuPhoi = item.GiaTriDieuPhoi || 0;
      const dieuChuyen = item.NhapDieuChuyen || 0;
      const giaTriDieuChuyen = item.GiaTriNhapDieuChuyen || 0;

      // Tổng xuất (bán) và tổng hủy
      const tongXuat = item.TongXuat || 0;
      const giaTriXuat = item.GiaTriXuat || 0;
      const tongHuy = item.TongHuy || 0;
      const giaTriHuy = item.GiaTriHuy || 0;

      // Tồn đầu kỳ
      const tonDK = item.TonDK || item.TonDauKy || 0;
      const giaTriDK = item.GiaTriDK || item.GiaTriDauKy || 0;

      // Tổng nhập
      const tongNhap = item.TongNhap || 0;
      const giaTriNhap = item.GiaTriNhap || 0;

      // Tồn cuối kỳ
      const tonCK = item.TonCK || 0;
      const giaTriCK = item.GiaTriCK || 0;

      const dataRow = worksheet.addRow([
        item.STT || index + 1,
        item.MaNL,
        item.Ten_SP,
        item.DVT_SP,
        tonCK > 0 ? giaTriCK / tonCK : tongNhap > 0 ? giaTriNhap / tongNhap : 0,
        tonDK,
        giaTriDK,
        tongNhap,
        giaTriNhap,
        dieuPhoi,
        giaTriDieuPhoi,
        dieuChuyen,
        giaTriDieuChuyen,
        tongXuat,
        giaTriXuat,
        tongHuy,
        giaTriHuy,
        tonCK,
        giaTriCK,
      ]);

      // Cộng dồn tổng
      totalTonDK += tonDK;
      totalGiaTriDK += giaTriDK;
      totalTongNhap += tongNhap;
      totalGiaTriNhap += giaTriNhap;
      totalDieuPhoi += dieuPhoi;
      totalGiaTriDieuPhoi += giaTriDieuPhoi;
      totalDieuChuyen += dieuChuyen;
      totalGiaTriDieuChuyen += giaTriDieuChuyen;
      totalTongXuat += tongXuat;
      totalGiaTriXuat += giaTriXuat;
      totalTongHuy += tongHuy;
      totalGiaTriHuy += giaTriHuy;
      totalTonCK += tonCK;
      totalGiaTriCK += giaTriCK;

      // Định dạng số và tiền tệ
      for (let col of [5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16, 17, 18, 19]) {
        dataRow.getCell(col).numFmt = "#,##0";
      }

      // Border cho dữ liệu
      for (let col = 1; col <= 19; col++) {
        dataRow.getCell(col).border = {
          top: { style: "thin" },
          left: { style: "thin" },
          bottom: { style: "thin" },
          right: { style: "thin" },
        };
      }

      // Căn lề
      dataRow.getCell(1).alignment = { horizontal: "center" };
      dataRow.getCell(2).alignment = { horizontal: "center" };
      dataRow.getCell(4).alignment = { horizontal: "center" };
      for (let col = 5; col <= 19; col++) {
        dataRow.getCell(col).alignment = { horizontal: "right" };
      }
    });

    // Thêm hàng tổng cộng
    const totalRow = worksheet.addRow([
      "",
      "",
      "TỔNG CỘNG",
      "",
      "",
      totalTonDK,
      totalGiaTriDK,
      totalTongNhap,
      totalGiaTriNhap,
      totalDieuPhoi,
      totalGiaTriDieuPhoi,
      totalDieuChuyen,
      totalGiaTriDieuChuyen,
      totalTongXuat,
      totalGiaTriXuat,
      totalTongHuy,
      totalGiaTriHuy,
      totalTonCK,
      totalGiaTriCK,
    ]);

    // Định dạng hàng tổng
    totalRow.font = { bold: true };
    totalRow.fill = {
      type: "pattern",
      pattern: "solid",
      fgColor: { argb: "FFFFCCCC" },
    };

    // Định dạng số cho hàng tổng
    for (let col = 6; col <= 19; col++) {
      totalRow.getCell(col).numFmt = "#,##0";
    }

    // Border cho hàng tổng
    for (let col = 1; col <= 19; col++) {
      totalRow.getCell(col).border = {
        top: { style: "thick" },
        left: { style: "thin" },
        bottom: { style: "thick" },
        right: { style: "thin" },
      };
    }

    // Căn lề hàng tổng
    totalRow.getCell(3).alignment = { horizontal: "center" };
    for (let col = 6; col <= 19; col++) {
      totalRow.getCell(col).alignment = { horizontal: "right" };
    }

    // Xuất file
    const buffer = await workbook.xlsx.writeBuffer();
    const blob = new Blob([buffer], {
      type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    });
    const url = window.URL.createObjectURL(blob);
    const link = document.createElement("a");
    link.href = url;
    link.download = `PhieuTonKho_${filters.tuNgay}_${filters.denNgay}.xlsx`;
    document.body.appendChild(link);
    link.click();
    document.body.removeChild(link);
    window.URL.revokeObjectURL(url);
  };

  // Xuất Kết Ca
  const exportKetCa = async () => {
    const workbook = new ExcelJS.Workbook();
    const worksheet = workbook.addWorksheet("Kết Ca");

    // Thêm tiêu đề lớn
    worksheet.mergeCells("A1:G1");
    worksheet.getCell("A1").value = "BÁO CÁO KẾT CA";
    worksheet.getCell("A1").font = { size: 16, bold: true };
    worksheet.getCell("A1").alignment = {
      vertical: "middle",
      horizontal: "center",
    };

    worksheet.mergeCells("A2:G2");
    worksheet.getCell(
      "A2"
    ).value = `Từ ngày ${filters.tuNgay} đến ngày ${filters.denNgay}`;
    worksheet.getCell("A2").font = { italic: true };
    worksheet.getCell("A2").alignment = {
      vertical: "middle",
      horizontal: "center",
    };

    worksheet.addRow([]);

    // Set column widths
    worksheet.getColumn(1).width = 15; // Ngày
    worksheet.getColumn(2).width = 10; // Ca
    worksheet.getColumn(3).width = 20; // Nhân viên
    worksheet.getColumn(4).width = 15; // Thời gian kết ca
    worksheet.getColumn(5).width = 18; // Tổng tiền
    worksheet.getColumn(6).width = 15; // Giảm giá
    worksheet.getColumn(7).width = 18; // Tiền sau giảm giá

    let totalTongTien = 0;
    let totalGiamGia = 0;
    let totalTienSauGiamGia = 0;
    let totalShifts = 0;

    shiftData.forEach((dayData) => {
      // Thêm header ngày
      const dateHeaderRow = worksheet.addRow([
        new Date(dayData.Ngay).toLocaleDateString("vi-VN"),
        "",
        "",
        "",
        "",
        "",
        "",
      ]);
      dateHeaderRow.font = { bold: true, size: 12 };
      dateHeaderRow.fill = {
        type: "pattern",
        pattern: "solid",
        fgColor: { argb: "FFE6F3FF" },
      };
      worksheet.mergeCells(`A${dateHeaderRow.number}:G${dateHeaderRow.number}`);

      // Thêm tiêu đề cột
      const headerRow = worksheet.addRow([
        "Ngày",
        "Ca",
        "Nhân Viên",
        "Thời Gian Kết Ca",
        "Tổng Tiền",
        "Giảm Giá",
        "Tiền Sau Giảm Giá",
      ]);
      headerRow.font = { bold: true };
      headerRow.alignment = { horizontal: "center" };
      headerRow.fill = {
        type: "pattern",
        pattern: "solid",
        fgColor: { argb: "FFDDEEFF" },
      };

      // Thêm dữ liệu ca
      dayData.Ca.forEach((shift) => {
        const tongTienCombined = shift.TongTien + (shift.DSBanhKem || 0);
        const tienSauGiamGiaCombined =
          shift.TienSauGiamGia + (shift.DSBanhKem || 0);

        worksheet.addRow([
          new Date(dayData.Ngay).toLocaleDateString("vi-VN"),
          `Ca ${shift.CaID}`,
          shift.TenNhanVien,
          shift.ThoiGianKetCa,
          tongTienCombined,
          shift.GiamGia,
          tienSauGiamGiaCombined,
        ]);
        totalTongTien += tongTienCombined;
        totalGiamGia += shift.GiamGia;
        totalTienSauGiamGia += tienSauGiamGiaCombined;
        totalShifts++;
      });

      // Thêm tổng ngày
      const dayTotalRow = worksheet.addRow([
        "",
        `Tổng ngày (${dayData.Ca.length} ca)`,
        "",
        "",
        dayData.Ca.reduce(
          (sum, shift) => sum + shift.TongTien + (shift.DSBanhKem || 0),
          0
        ),
        dayData.Ca.reduce((sum, shift) => sum + shift.GiamGia, 0),
        dayData.Ca.reduce(
          (sum, shift) => sum + shift.TienSauGiamGia + (shift.DSBanhKem || 0),
          0
        ),
      ]);
      dayTotalRow.font = { bold: true };
      dayTotalRow.fill = {
        type: "pattern",
        pattern: "solid",
        fgColor: { argb: "FFF0F8FF" },
      };

      worksheet.addRow([]); // Empty row between days
    });

    // Thêm hàng tổng cộng cuối
    const totalRow = worksheet.addRow([
      "",
      `TỔNG CỘNG (${totalShifts} ca)`,
      "",
      "",
      totalTongTien,
      totalGiamGia,
      totalTienSauGiamGia,
    ]);
    totalRow.font = { bold: true, size: 12 };
    totalRow.fill = {
      type: "pattern",
      pattern: "solid",
      fgColor: { argb: "FFFFCCCC" },
    };

    // Format numbers
    worksheet.eachRow((row, rowNumber) => {
      if (rowNumber > 4) {
        row.getCell(5).numFmt = "#,##0";
        row.getCell(6).numFmt = "#,##0";
        row.getCell(7).numFmt = "#,##0";
      }
    });

    const buffer = await workbook.xlsx.writeBuffer();
    const blob = new Blob([buffer], {
      type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    });
    const url = window.URL.createObjectURL(blob);
    const a = document.createElement("a");
    a.href = url;
    a.download = `BaoCaoKetCa_${filters.tuNgay}_${filters.denNgay}.xlsx`;
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
    window.URL.revokeObjectURL(url);
  };

  // Xuất Phiếu Kiểm Kê
  const exportKiemKe = async () => {
    if (kiemKeData.length === 0) {
      alert("Không có dữ liệu phiếu kiểm kê để xuất báo cáo");
      return;
    }

    const workbook = new ExcelJS.Workbook();

    // Worksheet tổng hợp
    const summaryWS = workbook.addWorksheet("Tổng hợp phiếu kiểm kê");

    // Thiết lập cột cho worksheet tổng hợp
    summaryWS.columns = [
      { header: "ID Phiếu", key: "IDPhieu", width: 12 },
      { header: "Ngày kiểm kê", key: "NgayKiemKe", width: 15 },
      { header: "Người kiểm kê", key: "TenUser", width: 20 },
      { header: "Ghi chú", key: "GhiChu", width: 30 },
      { header: "Tổng mặt hàng", key: "TongMatHang", width: 15 },
      { header: "Tổng chênh lệch SL", key: "TongChenhLech", width: 18 },
      { header: "Giá trị chênh lệch", key: "GiaTriChenhLech", width: 20 },
    ];

    // Thêm tiêu đề cho worksheet tổng hợp
    summaryWS.mergeCells("A1:G1");
    const titleCell = summaryWS.getCell("A1");
    titleCell.value = "BÁO CÁO TỔNG HỢP PHIẾU KIỂM KÊ";
    titleCell.font = { size: 16, bold: true };
    titleCell.alignment = { horizontal: "center", vertical: "middle" };

    summaryWS.mergeCells("A2:G2");
    const dateCell = summaryWS.getCell("A2");
    dateCell.value = `Từ ngày ${filters.tuNgay} đến ngày ${filters.denNgay}`;
    dateCell.font = { size: 12, italic: true };
    dateCell.alignment = { horizontal: "center", vertical: "middle" };

    summaryWS.addRow([]);

    // Header cho bảng tổng hợp
    const headerRow = summaryWS.addRow([
      "ID Phiếu",
      "Ngày kiểm kê",
      "Người kiểm kê",
      "Ghi chú",
      "Tổng mặt hàng",
      "Tổng chênh lệch SL",
      "Giá trị chênh lệch",
    ]);
    headerRow.font = { bold: true };
    headerRow.alignment = { horizontal: "center" };
    headerRow.fill = {
      type: "pattern",
      pattern: "solid",
      fgColor: { argb: "FFE6E6FA" },
    };

    let totalItems = 0;
    let totalQuantityDifference = 0;
    let totalValueDifference = 0;

    // Sort kiemKeData by date ascending
    const sortedKiemKeData = [...kiemKeData].sort(
      (a, b) => new Date(a.NgayKiemKe) - new Date(b.NgayKiemKe)
    );

    // Thêm dữ liệu cho từng phiếu kiểm kê
    for (const phieu of sortedKiemKeData) {
      try {
        // Fetch chi tiết cho từng phiếu
        const detailResponse = await fetch(
          `https://pos.doanquochoa.name.vn/api/chi-tiet-phieu-kiem-ke/${phieu.IDPhieu}`
        );
        let chiTietData = [];
        let tongMatHang = 0;
        let tongChenhLechSL = 0;
        let giaTriChenhLech = 0;

        if (detailResponse.ok) {
          chiTietData = await detailResponse.json();
          tongMatHang = chiTietData.length;

          chiTietData.forEach((item) => {
            const chenhLech = item.TonThucTe - item.TonSoSach;
            const thanhTienCL = chenhLech * item.DonGia;
            tongChenhLechSL += Math.abs(chenhLech);
            giaTriChenhLech += Math.abs(thanhTienCL);
          });
        }

        // Thêm dòng tổng hợp
        const dataRow = summaryWS.addRow([
          `#${phieu.IDPhieu}`,
          new Date(phieu.NgayKiemKe).toLocaleDateString("vi-VN"),
          phieu.TenUser,
          phieu.GhiChu || "-",
          tongMatHang,
          tongChenhLechSL,
          giaTriChenhLech,
        ]);

        // Format số
        dataRow.getCell(5).numFmt = "#,##0";
        dataRow.getCell(6).numFmt = "#,##0.00";
        dataRow.getCell(7).numFmt = "#,##0";

        totalItems += tongMatHang;
        totalQuantityDifference += tongChenhLechSL;
        totalValueDifference += giaTriChenhLech;

        // Tạo worksheet chi tiết cho từng phiếu
        if (chiTietData.length > 0) {
          const detailWS = workbook.addWorksheet(`Chi tiết #${phieu.IDPhieu}`);

          // Thiết lập cột
          detailWS.columns = [
            { header: "STT", key: "STT", width: 8 },
            { header: "Mã SP", key: "MaSP", width: 15 },
            { header: "Tên sản phẩm", key: "TenSP", width: 35 },
            { header: "ĐVT", key: "DonViTinh", width: 10 },
            { header: "Đơn giá", key: "DonGia", width: 15 },
            { header: "Tồn sổ sách", key: "TonSoSach", width: 15 },
            { header: "Tồn thực tế", key: "TonThucTe", width: 15 },
            { header: "Chênh lệch", key: "ChenhLech", width: 15 },
            { header: "Thành tiền CL", key: "ThanhTienCL", width: 18 },
          ];

          // Tiêu đề
          detailWS.mergeCells("A1:I1");
          const detailTitle = detailWS.getCell("A1");
          detailTitle.value = `CHI TIẾT PHIẾU KIỂM KÊ #${phieu.IDPhieu}`;
          detailTitle.font = { size: 14, bold: true };
          detailTitle.alignment = { horizontal: "center" };

          detailWS.mergeCells("A2:I2");
          const detailInfo = detailWS.getCell("A2");
          detailInfo.value = `Ngày: ${new Date(
            phieu.NgayKiemKe
          ).toLocaleDateString("vi-VN")} | Người kiểm kê: ${phieu.TenUser}`;
          detailInfo.font = { italic: true };
          detailInfo.alignment = { horizontal: "center" };

          if (phieu.GhiChu) {
            detailWS.mergeCells("A3:I3");
            const noteInfo = detailWS.getCell("A3");
            noteInfo.value = `Ghi chú: ${phieu.GhiChu}`;
            noteInfo.font = { italic: true };
            noteInfo.alignment = { horizontal: "left" };
          }

          detailWS.addRow([]);

          // Header
          const detailHeaderRow = detailWS.addRow([
            "STT",
            "Mã SP",
            "Tên sản phẩm",
            "ĐVT",
            "Đơn giá",
            "Tồn sổ sách",
            "Tồn thực tế",
            "Chênh lệch",
            "Thành tiền CL",
          ]);
          detailHeaderRow.font = { bold: true };
          detailHeaderRow.alignment = { horizontal: "center" };
          detailHeaderRow.fill = {
            type: "pattern",
            pattern: "solid",
            fgColor: { argb: "FFDDEEFF" },
          };

          // Dữ liệu chi tiết
          chiTietData.forEach((item, index) => {
            const chenhLech = item.TonThucTe - item.TonSoSach;
            const thanhTienCL = chenhLech * item.DonGia;

            const row = detailWS.addRow([
              index + 1,
              item.MaSP,
              item.TenSP,
              item.DonViTinh,
              item.DonGia,
              item.TonSoSach,
              item.TonThucTe,
              chenhLech,
              thanhTienCL,
            ]);

            // Format số
            row.getCell(5).numFmt = "#,##0";
            row.getCell(6).numFmt = "#,##0";
            row.getCell(7).numFmt = "#,##0";
            row.getCell(8).numFmt = "#,##0";
            row.getCell(9).numFmt = "#,##0";

            // Màu sắc cho chênh lệch
            if (chenhLech > 0) {
              row.getCell(8).font = { color: { argb: "FF10B981" }, bold: true };
              row.getCell(9).font = { color: { argb: "FF10B981" }, bold: true };
            } else if (chenhLech < 0) {
              row.getCell(8).font = { color: { argb: "FFEF4444" }, bold: true };
              row.getCell(9).font = { color: { argb: "FFEF4444" }, bold: true };
            }
          });

          // Tổng cộng cho chi tiết
          const detailTotalRow = detailWS.addRow([
            "",
            "",
            "TỔNG CỘNG",
            "",
            "",
            chiTietData.reduce((sum, item) => sum + item.TonSoSach, 0),
            chiTietData.reduce((sum, item) => sum + item.TonThucTe, 0),
            chiTietData.reduce(
              (sum, item) => sum + Math.abs(item.TonThucTe - item.TonSoSach),
              0
            ),
            chiTietData.reduce(
              (sum, item) =>
                sum + Math.abs((item.TonThucTe - item.TonSoSach) * item.DonGia),
              0
            ),
          ]);
          detailTotalRow.font = { bold: true };
          detailTotalRow.fill = {
            type: "pattern",
            pattern: "solid",
            fgColor: { argb: "FFFFCCCC" },
          };

          // Format total row numbers
          detailTotalRow.getCell(6).numFmt = "#,##0";
          detailTotalRow.getCell(7).numFmt = "#,##0";
          detailTotalRow.getCell(8).numFmt = "#,##0";
          detailTotalRow.getCell(9).numFmt = "#,##0";
        }
      } catch (error) {
        console.error(`Lỗi khi xử lý phiếu ${phieu.IDPhieu}:`, error);
      }
    }

    // Thêm tổng cộng cho worksheet tổng hợp
    const totalRow = summaryWS.addRow([
      "",
      "TỔNG CỘNG",
      "",
      "",
      totalItems,
      totalQuantityDifference,
      totalValueDifference,
    ]);
    totalRow.font = { bold: true };
    totalRow.fill = {
      type: "pattern",
      pattern: "solid",
      fgColor: { argb: "FFFFD700" },
    };

    // Format total row
    totalRow.getCell(5).numFmt = "#,##0";
    totalRow.getCell(6).numFmt = "#,##0.00";
    totalRow.getCell(7).numFmt = "#,##0";

    // Xuất file
    const buffer = await workbook.xlsx.writeBuffer();
    const blob = new Blob([buffer], {
      type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    });
    const url = window.URL.createObjectURL(blob);
    const link = document.createElement("a");
    link.href = url;
    link.download = `PhieuKiemKe_${filters.tuNgay}_${filters.denNgay}.xlsx`;
    document.body.appendChild(link);
    link.click();
    document.body.removeChild(link);
    window.URL.revokeObjectURL(url);
  };

  // Xuất Bánh Kem
  const exportBanhKem = async () => {
    const workbook = new ExcelJS.Workbook();
    const worksheet = workbook.addWorksheet("Chi Tiết Bán Bánh Kem");

    if (banhKemData.length === 0) {
      worksheet.addRow(["Không có dữ liệu chi tiết bán bánh kem"]);
    } else {
      // Thêm tiêu đề lớn
      worksheet.mergeCells("A1:J1");
      worksheet.getCell("A1").value = "BÁO CÁO CHI TIẾT BÁN BÁNH KEM";
      worksheet.getCell("A1").font = { size: 16, bold: true };
      worksheet.getCell("A1").alignment = {
        vertical: "middle",
        horizontal: "center",
      };

      worksheet.mergeCells("A2:J2");
      worksheet.getCell(
        "A2"
      ).value = `Từ ngày ${filters.tuNgay} đến ngày ${filters.denNgay}`;
      worksheet.getCell("A2").font = { italic: true };
      worksheet.getCell("A2").alignment = {
        vertical: "middle",
        horizontal: "center",
      };

      // Thêm thông tin cửa hàng được chọn
      const selectedStoreNames = stores
        .filter((store) => filters.dsMaCH.includes(store.MaCuaHang))
        .map((store) => store.TenCuaHang)
        .join(", ");

      worksheet.mergeCells("A3:J3");
      worksheet.getCell("A3").value = `Cửa hàng: ${selectedStoreNames}`;
      worksheet.getCell("A3").font = { italic: true };
      worksheet.getCell("A3").alignment = {
        vertical: "middle",
        horizontal: "center",
      };

      worksheet.addRow([]);

      // Thêm tiêu đề cột
      const headerRow = worksheet.addRow([
        "ID Điều Chuyển",
        "Sale Order",
        "Mã SP",
        "Tên SP",
        "Số Lượng",
        "Đơn Giá",
        "Thành Tiền",
        "Ngày Thực Hiện",
        "Nhân Viên",
        "Cửa Hàng",
      ]);
      headerRow.font = { bold: true };
      headerRow.alignment = { horizontal: "center" };
      headerRow.fill = {
        type: "pattern",
        pattern: "solid",
        fgColor: { argb: "FFDDEEFF" },
      };

      // Set column widths
      worksheet.getColumn(1).width = 15;
      worksheet.getColumn(2).width = 15;
      worksheet.getColumn(3).width = 20;
      worksheet.getColumn(4).width = 30;
      worksheet.getColumn(5).width = 12;
      worksheet.getColumn(6).width = 15;
      worksheet.getColumn(7).width = 15;
      worksheet.getColumn(8).width = 18;
      worksheet.getColumn(9).width = 15;
      worksheet.getColumn(10).width = 12;

      let totalSoLuong = 0;
      let totalThanhTien = 0;

      // Sort data by date descending
      const sortedData = banhKemData.sort(
        (a, b) => new Date(b.NgayThucHien) - new Date(a.NgayThucHien)
      );

      sortedData.forEach((item) => {
        const thanhTien = item.SoLuong * item.DonGia;
        const storeName =
          stores.find((store) => store.MaCuaHang === item.MaCH)?.TenCuaHang ||
          item.MaCH;
        const localDate = new Date(
          new Date(item.NgayThucHien).getTime() - 7 * 60 * 60 * 1000
        );
        worksheet.addRow([
          item.IDDieuChuyen,
          item.Sale_Order,
          item.MaSP,
          item.TenSP,
          item.SoLuong,
          item.DonGia,
          thanhTien,
          localDate.toLocaleDateString("vi-VN") +
            " " +
            localDate.toLocaleTimeString("vi-VN", {
              hour: "2-digit",
              minute: "2-digit",
            }),
          item.NhanVien,
          storeName,
        ]);
        totalSoLuong += item.SoLuong;
        totalThanhTien += thanhTien;
      });

      // Thêm hàng tổng cộng
      const totalRow = worksheet.addRow([
        "",
        "",
        "",
        "TỔNG CỘNG",
        totalSoLuong,
        "",
        totalThanhTien,
        "",
        "",
        "",
      ]);
      totalRow.font = { bold: true };
      totalRow.fill = {
        type: "pattern",
        pattern: "solid",
        fgColor: { argb: "FFFFCCCC" },
      };

      // Format numbers
      worksheet.eachRow((row, rowNumber) => {
        if (rowNumber > 5) {
          row.getCell(5).numFmt = "#,##0";
          row.getCell(6).numFmt = "#,##0";
          row.getCell(7).numFmt = "#,##0";
        }
      });
    }

    const buffer = await workbook.xlsx.writeBuffer();
    const blob = new Blob([buffer], {
      type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    });
    const url = window.URL.createObjectURL(blob);
    const a = document.createElement("a");
    a.href = url;
    a.download = `ChiTietBanBanhKem_${filters.tuNgay}_${filters.denNgay}.xlsx`;
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
    window.URL.revokeObjectURL(url);
  };

  // Xuất Bánh Kem Đặt
  const exportWholesale = async () => {
    const workbook = new ExcelJS.Workbook();
    const worksheet = workbook.addWorksheet("Chi Tiết Bánh Kem Đặt");

    if (wholesaleData.length === 0) {
      worksheet.addRow(["Không có dữ liệu chi tiết bánh kem đặt"]);
    } else {
      // Thêm tiêu đề lớn
      worksheet.mergeCells("A1:J1");
      worksheet.getCell("A1").value = "BÁO CÁO CHI TIẾT BÁNH KEM ĐẶT";
      worksheet.getCell("A1").font = { size: 16, bold: true };
      worksheet.getCell("A1").alignment = {
        vertical: "middle",
        horizontal: "center",
      };

      worksheet.mergeCells("A2:J2");
      worksheet.getCell(
        "A2"
      ).value = `Từ ngày ${filters.tuNgay} đến ngày ${filters.denNgay}`;
      worksheet.getCell("A2").font = { italic: true };
      worksheet.getCell("A2").alignment = {
        vertical: "middle",
        horizontal: "center",
      };

      // Thêm thông tin cửa hàng được chọn
      const selectedStoreNames = stores
        .filter((store) => filters.dsMaCH.includes(store.IDCH))
        .map((store) => store.TenCuaHang)
        .join(", ");

      worksheet.mergeCells("A3:J3");
      worksheet.getCell("A3").value = `Cửa hàng: ${selectedStoreNames}`;
      worksheet.getCell("A3").font = { italic: true };
      worksheet.getCell("A3").alignment = {
        vertical: "middle",
        horizontal: "center",
      };

      worksheet.addRow([]);

      // Thêm tiêu đề cột
      const headerRow = worksheet.addRow([
        "Mã PB",
        "ID Cửa Hàng",
        "Mã SP",
        "Đơn Giá",
        "Số Lượng",
        "Thành Tiền",
        "Ngày Đặt",
        "Ngày Nhận",
        "Người Đặt",
        "Người Xác Nhận",
      ]);
      headerRow.font = { bold: true };
      headerRow.alignment = { horizontal: "center" };
      headerRow.fill = {
        type: "pattern",
        pattern: "solid",
        fgColor: { argb: "FFDDEEFF" },
      };

      // Set column widths
      worksheet.getColumn(1).width = 12;
      worksheet.getColumn(2).width = 12;
      worksheet.getColumn(3).width = 20;
      worksheet.getColumn(4).width = 15;
      worksheet.getColumn(5).width = 12;
      worksheet.getColumn(6).width = 15;
      worksheet.getColumn(7).width = 18;
      worksheet.getColumn(8).width = 18;
      worksheet.getColumn(9).width = 15;
      worksheet.getColumn(10).width = 15;

      let totalSoLuong = 0;
      let totalThanhTien = 0;

      // Sort data by date descending
      const sortedData = wholesaleData.sort(
        (a, b) => new Date(b.NgayDat) - new Date(a.NgayDat)
      );

      sortedData.forEach((item) => {
        const thanhTien = item.SoLuong * item.DonGia;
        const ngayDat = new Date(item.NgayDat);
        const ngayNhan = new Date(item.NgayNhan);

        worksheet.addRow([
          item.MaPB,
          item.IDCH,
          item.MaSP,
          item.DonGia,
          item.SoLuong,
          thanhTien,
          ngayDat.toLocaleDateString("vi-VN") +
            " " +
            ngayDat.toLocaleTimeString("vi-VN", {
              hour: "2-digit",
              minute: "2-digit",
            }),
          ngayNhan.toLocaleDateString("vi-VN") +
            " " +
            ngayNhan.toLocaleTimeString("vi-VN", {
              hour: "2-digit",
              minute: "2-digit",
            }),
          item.NguoiDat,
          item.NguoiXacNhan,
        ]);
        totalSoLuong += item.SoLuong;
        totalThanhTien += thanhTien;
      });

      // Thêm hàng tổng cộng
      const totalRow = worksheet.addRow([
        "",
        "",
        "TỔNG CỘNG",
        "",
        totalSoLuong,
        "",
        totalThanhTien,
        "",
        "",
        "",
      ]);
      totalRow.font = { bold: true };
      totalRow.fill = {
        type: "pattern",
        pattern: "solid",
        fgColor: { argb: "FFFFCCCC" },
      };

      // Format numbers
      worksheet.eachRow((row, rowNumber) => {
        if (rowNumber > 5) {
          row.getCell(4).numFmt = "#,##0";
          row.getCell(5).numFmt = "#,##0";
          row.getCell(6).numFmt = "#,##0";
        }
      });
    }

    const buffer = await workbook.xlsx.writeBuffer();
    const blob = new Blob([buffer], {
      type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    });
    const url = window.URL.createObjectURL(blob);
    const a = document.createElement("a");
    a.href = url;
    a.download = `ChiTietBanhKemDat_${filters.tuNgay}_${filters.denNgay}.xlsx`;
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
    window.URL.revokeObjectURL(url);
  };

  // Map export key -> loading state key
  const exportLoadingMap = {
    doanhSo: "thongKe",
    doanhSoSP: "thongKe",
    hoaDon: "thongKe",
    chiTietBanHang: "thongKe",
    nhapHang: "nhapHang",
    tonKho: "khoHang",
    ketCa: "ketCa",
    kiemKe: "kiemKe",
    banhKem: "banhKem",
    wholesale: "wholesale",
  };

  // Kiểm tra 1 export option có đang loading không
  const isExportLoading = (key) => {
    const loadingKey = exportLoadingMap[key];
    return loadingKey ? loadingState[loadingKey] : false;
  };

  const exportToExcel = async () => {
    const selectedOptions = Object.entries(exportOptions).filter(
      ([key, value]) => value
    );

    if (selectedOptions.length === 0) {
      alert("Vui lòng chọn ít nhất một loại báo cáo để xuất");
      return;
    }

    // Kiểm tra xem có option nào đang loading không
    const stillLoading = selectedOptions.filter(([key]) => isExportLoading(key));
    if (stillLoading.length > 0) {
      alert("Một số dữ liệu vẫn đang tải. Vui lòng đợi hoặc bỏ chọn các mục đang tải.");
      return;
    }

    // Nếu chọn xuất tồn kho, fetch dữ liệu kho hàng trước
    if (exportOptions.tonKho && khoHangData.length === 0) {
      await fetchKhoHang();
    }

    // Sử dụng Promise.all để chờ tất cả các file xuất xong
    try {
      await Promise.all(
        selectedOptions.map(([option]) => {
          switch (option) {
            case "doanhSo":
              return exportDoanhSo();
            case "doanhSoSP":
              return exportDoanhSoSP();
            case "hoaDon":
              return exportHoaDon();
            case "chiTietBanHang":
              return exportChiTietBanHang();
            case "nhapHang":
              return exportNhapHang();
            case "tonKho":
              return exportTonKho();
            case "ketCa":
              return exportKetCa();
            case "kiemKe":
              return exportKiemKe();
            case "banhKem":
              return exportBanhKem();
            case "wholesale":
              return exportWholesale();
            default:
              return Promise.resolve();
          }
        })
      );
      const fileCount = selectedOptions.length;
      alert(`Đã xuất ${fileCount} file Excel thành công!`);
    } catch (error) {
      console.error("Lỗi khi xuất báo cáo:", error);
      alert("Có lỗi xảy ra khi xuất báo cáo. Vui lòng thử lại.");
    }
  };

  // Chart data and options
  const revenueChartData = {
    labels: data.THDoanhSo.map((item) => item.NgayThangNam),
    datasets: [
      {
        label: "Doanh Thu Tổng Hợp",
        data: data.THDoanhSo.map((item) => {
          // Tính doanh thu tổng hợp cho từng ngày
          const ngay = item.NgayThangNam;
          const banhKemRevenueForDay = banhKemData
            .filter(bk => new Date(bk.NgayThucHien).toDateString() === new Date(ngay).toDateString())
            .reduce((sum, bk) => sum + (bk.SoLuong * bk.DonGia), 0);
          const wholesaleRevenueForDay = wholesaleData
            .filter(ws => new Date(ws.NgayDat).toDateString() === new Date(ngay).toDateString())
            .reduce((sum, ws) => sum + (ws.SoLuong * ws.DonGia), 0);
          return item.DoanhThu + banhKemRevenueForDay + wholesaleRevenueForDay;
        }),
        backgroundColor: "rgba(54, 162, 235, 0.6)",
        borderColor: "rgba(54, 162, 235, 1)",
        borderWidth: 1,
      },
      {
        label: "Doanh Thu Sau Giảm Giá",
        data: data.THDoanhSo.map((item) => {
          // Tính doanh thu sau giảm giá tổng hợp cho từng ngày
          const ngay = item.NgayThangNam;
          const banhKemRevenueForDay = banhKemData
            .filter(bk => new Date(bk.NgayThucHien).toDateString() === new Date(ngay).toDateString())
            .reduce((sum, bk) => sum + (bk.SoLuong * bk.DonGia), 0);
          const wholesaleRevenueForDay = wholesaleData
            .filter(ws => new Date(ws.NgayDat).toDateString() === new Date(ngay).toDateString())
            .reduce((sum, ws) => sum + (ws.SoLuong * ws.DonGia), 0);
          return item.DoanhThuConLai + banhKemRevenueForDay + wholesaleRevenueForDay;
        }),
        backgroundColor: "rgba(255, 99, 132, 0.6)",
        borderColor: "rgba(255, 99, 132, 1)",
        borderWidth: 1,
      },
    ],
  };

  // Biểu đồ top sản phẩm bán chạy
  const topProducts = data.THDoanhSoSP.sort(
    (a, b) => b.SoLuong - a.SoLuong
  ).slice(0, 10);
  const productChartData = {
    labels: topProducts.map((item) => item.TenSP),
    datasets: [
      {
        label: "Số Lượng Bán",
        data: topProducts.map((item) => item.SoLuong),
        backgroundColor: [
          "#FF6384",
          "#36A2EB",
          "#FFCE56",
          "#4BC0C0",
          "#9966FF",
          "#FF9F40",
          "#FF6384",
          "#C9CBCF",
          "#4BC0C0",
          "#FF6384",
        ],
      },
    ],
  };

  // Biểu đồ doanh thu theo cửa hàng
  const storeRevenue = data.THDoanhSo.reduce((acc, item) => {
    if (!acc[item.TenCuaHang]) {
      acc[item.TenCuaHang] = 0;
    }
    acc[item.TenCuaHang] += item.DoanhThuConLai;
    return acc;
  }, {});

  // Thêm doanh thu bánh kem theo cửa hàng
  banhKemData.forEach(item => {
    const storeName = stores.find(store => store.MaCuaHang === item.MaCH)?.TenCuaHang || item.MaCuaHang;
    if (!storeRevenue[storeName]) {
      storeRevenue[storeName] = 0;
    }
    storeRevenue[storeName] += item.SoLuong * item.DonGia;
  });

  // Thêm doanh thu bánh kem đặt theo cửa hàng
  wholesaleData.forEach(item => {
    const storeName = stores.find(store => store.IDCH === item.IDCH)?.TenCuaHang || item.IDCH;
    if (!storeRevenue[storeName]) {
      storeRevenue[storeName] = 0;
    }
    storeRevenue[storeName] += item.SoLuong * item.DonGia;
  });

  const storeChartData = {
    labels: Object.keys(storeRevenue),
    datasets: [
      {
        data: Object.values(storeRevenue),
        backgroundColor: [
          "#FF6384",
          "#36A2EB",
          "#FFCE56",
          "#4BC0C0",
          "#9966FF",
        ],
      },
    ],
  };

  const formatCurrency = (amount) => {
    return new Intl.NumberFormat("vi-VN", {
      style: "currency",
      currency: "VND",
      minimumFractionDigits: 0,
      maximumFractionDigits: 0,
    }).format(amount);
  };

  const formatNumber = (number) => {
    return new Intl.NumberFormat("vi-VN").format(number);
  };

  // Chart configurations with modern styling
  const chartOptions = {
    responsive: true,
    maintainAspectRatio: false,
    plugins: {
      legend: {
        position: "top",
        labels: {
          usePointStyle: true,
          padding: 20,
          font: {
            size: 12,
            family: "'Inter', sans-serif",
          },
        },
      },
      title: {
        display: false,
      },
    },
    scales: {
      y: {
        beginAtZero: true,
        grid: {
          color: "#f1f3f4",
          drawBorder: false,
        },
        ticks: {
          font: {
            size: 11,
            family: "'Inter', sans-serif",
          },
          color: "#6b7280",
        },
      },
      x: {
        grid: {
          display: false,
        },
        ticks: {
          font: {
            size: 11,
            family: "'Inter', sans-serif",
          },
          color: "#6b7280",
        },
      },
    },
  };

  const pieOptions = {
    responsive: true,
    maintainAspectRatio: false,
    plugins: {
      legend: {
        position: "bottom",
        labels: {
          usePointStyle: true,
          padding: 15,
          font: {
            size: 11,
            family: "'Inter', sans-serif",
          },
        },
      },
    },
  };

  // Calculate summary statistics
  const totalRevenue = data.THDoanhSo.reduce(
    (sum, item) => sum + item.DoanhThu,
    0
  );
  const totalBanhKemRevenue = banhKemData.reduce(
    (sum, item) => sum + item.SoLuong * item.DonGia,
    0
  );
  const totalWholesaleRevenue = wholesaleData.reduce(
    (sum, item) => sum + item.SoLuong * item.DonGia,
    0
  );
  const totalCombinedRevenue =
    totalRevenue + totalWholesaleRevenue; // Tổng doanh thu kết hợp
  const totalOrders = data.THHoaDon.length;
  const totalProducts = data.THDoanhSoSP.reduce(
    (sum, item) => sum + item.SoLuong,
    0
  );
  const totalBanhKemProducts = banhKemData.reduce(
    (sum, item) => sum + item.SoLuong,
    0
  );
  const totalWholesaleProducts = wholesaleData.reduce(
    (sum, item) => sum + item.SoLuong,
    0
  );
  const totalDiscount = data.THDoanhSo.reduce(
    (sum, item) => sum + item.GiamGia,
    0
  );
  const totalInventoryValue = khoHangData.reduce(
    (sum, item) => sum + (item.GiaTriCK || 0),
    0
  );
  const totalInventoryItems = khoHangData.reduce(
    (sum, item) => sum + (item.TonCK || 0),
    0
  );

  // Shift statistics
  const totalShifts = shiftData.reduce((sum, day) => sum + day.Ca.length, 0);
  const totalShiftRevenue = shiftData.reduce(
    (sum, day) =>
      sum +
      day.Ca.reduce(
        (daySum, shift) =>
          daySum + (shift.TienSauGiamGia + (shift.DSBanhKem || 0)),
        0
      ),
    0
  );

  // Inventory check statistics
  const totalKiemKeRecords = kiemKeData.length;

  // Filter stores based on search
  const filteredStores = stores.filter((store) =>
    store.TenCuaHang.toLowerCase().includes(storeSearch.toLowerCase())
  );

  // Hàm lấy dữ liệu điều chuyển cho sản phẩm (nếu có)
  const getTransferDataForProduct = (item) => {
    // Nếu item có trường NhapDieuChuyen/GiaTriNhapDieuChuyen thì dùng, nếu không thì trả về 0
    return {
      nhapDieuChuyen: item.NhapDieuChuyen || 0,
      giaTriNhapDieuChuyen: item.GiaTriNhapDieuChuyen || 0,
    };
  };

  return (
    <div className="dashboard-container">
      {/* Header */}
      <div className="dashboard-header">
        <div className="header-content">
          <div className="header-title">
            <div className="title-icon">
              <FiBarChart2 />
            </div>
            <div className="title-text">
              <h1>Báo Cáo Kinh Doanh</h1>
              <p>Tổng quan và phân tích dữ liệu kinh doanh chi tiết</p>
            </div>
          </div>
          <div className="header-actions">
            <div className="export-dropdown">
              <div className="export-trigger" onClick={toggleExportDropdown}>
                <FiDownload />
                <span>Xuất báo cáo</span>
              </div>
              {showExportDropdown && (
                <div className="export-menu">
                  <div className="export-header">
                    <label className="export-all">
                      <span className="checkbox-icon">
                        {Object.values(exportOptions).every((v) => v) ? (
                          <FiCheckSquare />
                        ) : (
                          <FiSquare />
                        )}
                      </span>
                      <input
                        type="checkbox"
                        checked={Object.values(exportOptions).every((v) => v)}
                        onChange={(e) =>
                          handleSelectAllExports(e.target.checked)
                        }
                      />
                      Chọn tất cả
                    </label>
                  </div>
                  <div className="export-options-list">
                    {Object.entries({
                      doanhSo: "Doanh số",
                      doanhSoSP: "Doanh số sản phẩm",
                      hoaDon: "Hóa đơn",
                      chiTietBanHang: "Chi tiết bán hàng",
                      nhapHang: "Nhập hàng",
                      tonKho: "Phiếu tồn kho",
                      ketCa: "Báo cáo kết ca",
                      kiemKe: "Phiếu kiểm kê",
                      banhKem: "Chi tiết bán bánh kem",
                      wholesale: "Chi tiết bánh kem đặt",
                    }).map(([key, label]) => {
                      const itemLoading = isExportLoading(key);
                      return (
                        <label key={key} className={`export-option ${itemLoading ? "export-option-loading" : ""}`}
                          style={itemLoading ? { opacity: 0.5, pointerEvents: "none" } : {}}
                        >
                          <span className="option-icon">
                            {itemLoading ? <FiRefreshCw className="spinning" /> : <FiFileText />}
                          </span>
                          <span className="option-label">
                            {label}
                            {itemLoading && <span style={{ fontSize: "0.75em", marginLeft: 6, color: "#888" }}>(đang tải...)</span>}
                          </span>
                          <span className="checkbox-icon">
                            {exportOptions[key] ? (
                              <FiCheckSquare />
                            ) : (
                              <FiSquare />
                            )}
                          </span>
                          <input
                            type="checkbox"
                            checked={exportOptions[key]}
                            onChange={(e) =>
                              handleExportOptionChange(key, e.target.checked)
                            }
                            disabled={itemLoading}
                          />
                        </label>
                      );
                    })}
                  </div>
                  <button
                    className="export-button"
                    onClick={() => {
                      exportToExcel();
                      setShowExportDropdown(false);
                    }}
                    disabled={Object.entries(exportOptions).some(([key, val]) => val && isExportLoading(key))}
                  >
                    <FiDownload />
                    {Object.entries(exportOptions).some(([key, val]) => val && isExportLoading(key))
                      ? "Đang tải dữ liệu..."
                      : "Xuất Excel"}
                  </button>
                </div>
              )}
            </div>
          </div>
        </div>
      </div>

      {/* Filters */}
      <div className="filters-section">
        <div className="filters-card">
          <div className="filters-header">
            <div className="filters-title">
              <FiFilter />
              <h3>Bộ lọc dữ liệu</h3>
            </div>
            <button
              className="refresh-btn"
              onClick={fetchData}
              disabled={isAnyLoading}
            >
              <FiRefreshCw className={isAnyLoading ? "spinning" : ""} />
              {isAnyLoading ? "Đang tải..." : "Cập nhật"}
            </button>
          </div>
          <div className="filters-content">
            <div className="filter-group">
              <label>
                <FiCalendar />
                Từ ngày
              </label>
              <input
                type="date"
                value={filters.tuNgay}
                onChange={(e) => handleFilterChange("tuNgay", e.target.value)}
                className="date-input"
              />
            </div>
            <div className="filter-group">
              <label>
                <FiCalendar />
                Đến ngày
              </label>
              <input
                type="date"
                value={filters.denNgay}
                onChange={(e) => handleFilterChange("denNgay", e.target.value)}
                className="date-input"
              />
            </div>
            <div className="filter-group stores-filter">
              <label>
                <FiShoppingCart />
                Cửa hàng ({filters.dsMaCH.length}/{stores.length})
              </label>
              <div className="stores-dropdown">
                <div className="stores-search">
                  <div className="search-input">
                    <FiSearch />
                    <input
                      type="text"
                      placeholder="Tìm kiếm cửa hàng..."
                      value={storeSearch}
                      onChange={(e) => setStoreSearch(e.target.value)}
                    />
                  </div>
                </div>
                <div className="stores-header">
                  <label className="select-all-stores">
                    <span className="checkbox-icon">
                      {filters.dsMaCH.length === stores.length ? (
                        <FiCheckSquare />
                      ) : (
                        <FiSquare />
                      )}
                    </span>
                    <input
                      type="checkbox"
                      checked={filters.dsMaCH.length === stores.length}
                      onChange={(e) => handleSelectAllStores(e.target.checked)}
                    />
                    Chọn tất cả
                  </label>
                </div>
                <div className="stores-list">
                  {filteredStores.map((store) => (
                    <label key={store.IDCH} className="store-option">
                      <span className="checkbox-icon">
                        {filters.dsMaCH.includes(store.IDCH) ? (
                          <FiCheckSquare />
                        ) : (
                          <FiSquare />
                        )}
                      </span>
                      <input
                        type="checkbox"
                        checked={filters.dsMaCH.includes(store.IDCH)}
                        onChange={(e) =>
                          handleStoreChange(store.IDCH, e.target.checked)
                        }
                      />
                      <span className="store-name">{store.TenCuaHang}</span>
                    </label>
                  ))}
                  {filteredStores.length === 0 && storeSearch && (
                    <div className="no-results">
                      Không tìm thấy cửa hàng nào
                    </div>
                  )}
                </div>
              </div>
            </div>
          </div>
        </div>
      </div>

        <>
          {/* Summary Cards */}
          <div className="summary-section">
            <div className="summary-grid">
              <div className="summary-card revenue">
                <div className="card-icon">
                  <FiDollarSign />
                </div>
                <div className="card-content">
                  <div className="card-value">
                    {formatCurrency(totalCombinedRevenue)}
                  </div>
                  <div className="card-label">Tổng Doanh Thu</div>
                </div>
                <div className="card-trend">
                  <FiTrendingUp />
                </div>
              </div>
              <div className="summary-card orders">
                <div className="card-icon">
                  <FiFileText />
                </div>
                <div className="card-content">
                  <div className="card-value">{formatNumber(totalOrders)}</div>
                  <div className="card-label">Tổng Hóa Đơn</div>
                </div>
                <div className="card-trend">
                  <FiTrendingUp />
                </div>
              </div>
              <div className="summary-card products">
                <div className="card-icon">
                  <FiPackage />
                </div>
                <div className="card-content">
                  <div className="card-value">
                    {formatNumber(
                      totalProducts +
                        totalBanhKemProducts +
                        totalWholesaleProducts
                    )}
                  </div>
                  <div className="card-label">Sản Phẩm Bán</div>
                </div>
                <div className="card-trend">
                  <FiTrendingUp />
                </div>
              </div>
              <div className="summary-card discount">
                <div className="card-icon">
                  <FiCreditCard />
                </div>
                <div className="card-content">
                  <div className="card-value">
                    {formatCurrency(totalDiscount)}
                  </div>
                  <div className="card-label">Tổng Giảm Giá</div>
                </div>
                <div className="card-trend">
                  <FiTrendingUp />
                </div>
              </div>
              <div className="summary-card cake-revenue">
                <div className="card-icon">
                  <FiDollarSign />
                </div>
                <div className="card-content">
                  <div className="card-value">
                    {formatCurrency(totalBanhKemRevenue)}
                  </div>
                  <div className="card-label">Doanh Thu Bánh Kem Bán</div>
                </div>
                <div className="card-trend">
                  <FiTrendingUp />
                </div>
              </div>
              <div className="summary-card wholesale-revenue">
                <div className="card-icon">
                  <FiDollarSign />
                </div>
                <div className="card-content">
                  <div className="card-value">
                    {formatCurrency(totalWholesaleRevenue)}
                  </div>
                  <div className="card-label">Doanh Thu Bánh Kem Đặt</div>
                </div>
                <div className="card-trend">
                  <FiTrendingUp />
                </div>
              </div>
              <div className="summary-card other-revenue">
                <div className="card-icon">
                  <FiDollarSign />
                </div>
                <div className="card-content">
                  <div className="card-value">
                    {formatCurrency(totalRevenue)}
                  </div>
                  <div className="card-label">
                    Doanh Thu Bánh Mặn Ngọt, BMCN
                  </div>
                </div>
                <div className="card-trend">
                  <FiTrendingUp />
                </div>
              </div>
              <div className="summary-card inventory">
                <div className="card-icon">
                  <FiClipboard />
                </div>
                <div className="card-content">
                  <div className="card-value">
                    {formatNumber(totalKiemKeRecords)}
                  </div>
                  <div className="card-label">Phiếu Kiểm Kê</div>
                </div>
                <div className="card-change up">Trong kỳ</div>
              </div>
            </div>
          </div>

          {/* Charts */}
          <div className="charts-section">
            <div className="charts-grid">
              <div className="chart-card large">
                <div className="chart-header">
                  <div className="chart-title">
                    <FiBarChart2 />
                    <h3>Doanh Thu Theo Ngày</h3>
                  </div>
                  <div className="chart-actions">
                    <button className="chart-action">
                      <FiEye />
                    </button>
                  </div>
                </div>
                <div className="chart-content">
                  <Bar data={revenueChartData} options={chartOptions} />
                </div>
              </div>
              <div className="chart-card">
                <div className="chart-header">
                  <div className="chart-title">
                    <FiPackage />
                    <h3>Top 10 Sản Phẩm</h3>
                  </div>
                </div>
                <div className="chart-content">
                  <Bar data={productChartData} options={chartOptions} />
                </div>
              </div>
              <div className="chart-card">
                <div className="chart-header">
                  <div className="chart-title">
                    <FiPieChart />
                    <h3>Doanh Thu Cửa Hàng</h3>
                  </div>
                </div>
                <div className="chart-content">
                  <Pie data={storeChartData} options={pieOptions} />
                </div>
              </div>
            </div>
          </div>

          {/* Data Tables */}
          <div className="tables-section">
            <div className="table-card">
              <div className="table-header">
                <div className="table-title">
                  <FiDollarSign />
                  <h3>Thống Kê Doanh Số {loadingState.thongKe && <FiRefreshCw className="spinning" style={{ marginLeft: 8, fontSize: 14 }} />}</h3>
                </div>
                <div className="table-actions">
                  <button className="table-action">
                    <FiEye />
                    Xem chi tiết
                  </button>
                </div>
              </div>
              <div className="table-content">
                <div className="modern-table">
                  <table>
                    <thead>
                      <tr>
                        <th>Ngày</th>
                        <th>Cửa Hàng</th>
                        <th>Doanh Thu</th>
                        <th>Giảm Giá</th>
                        <th>Còn Lại</th>
                      </tr>
                    </thead>
                    <tbody>
                      {data.THDoanhSo.slice(0, 10).map((item, index) => {
                        // Tính doanh thu tổng hợp cho từng dòng
                        const ngay = item.NgayThangNam;
                        const tenCuaHang = item.TenCuaHang;
                        
                        const banhKemRevenueForRow = banhKemData
                          .filter(bk => 
                            new Date(bk.NgayThucHien).toDateString() === new Date(ngay).toDateString() &&
                            (stores.find(store => store.MaCuaHang === bk.MaCH)?.TenCuaHang || bk.MaCH) === tenCuaHang
                          )
                          .reduce((sum, bk) => sum + (bk.SoLuong * bk.DonGia), 0);
                        
                        const wholesaleRevenueForRow = wholesaleData
                          .filter(ws => 
                            new Date(ws.NgayDat).toDateString() === new Date(ngay).toDateString() &&
                            (stores.find(store => store.IDCH === ws.IDCH)?.TenCuaHang || ws.IDCH) === tenCuaHang
                          )
                          .reduce((sum, ws) => sum + (ws.SoLuong * ws.DonGia), 0);
                        
                        const totalRevenueForRow = item.DoanhThu + banhKemRevenueForRow + wholesaleRevenueForRow;
                        const totalNetRevenueForRow = item.DoanhThuConLai + banhKemRevenueForRow + wholesaleRevenueForRow;
                        
                        return (
                          <tr key={index}>
                            <td>
                              {new Date(item.NgayThangNam).toLocaleDateString(
                                "vi-VN"
                              )}
                            </td>
                            <td>
                              <div className="store-cell">
                                <FiShoppingCart />
                                {item.TenCuaHang}
                              </div>
                            </td>
                            <td className="amount">
                              {formatCurrency(totalRevenueForRow)}
                            </td>
                            <td className="discount">
                              {formatCurrency(item.GiamGia)}
                            </td>
                            <td className="net-amount">
                              {formatCurrency(totalNetRevenueForRow)}
                            </td>
                          </tr>
                        );
                      })}
                    </tbody>
                  </table>
                </div>
              </div>
            </div>

            {/* Bánh Kem Data Table */}
            <div className="table-card">
              <div className="table-header">
                <div className="table-title">
                  <FiPackage />
                  <h3>Chi Tiết Bán Bánh Kem {loadingState.banhKem && <FiRefreshCw className="spinning" style={{ marginLeft: 8, fontSize: 14 }} />}</h3>
                </div>
                <div className="table-actions">
                  <button
                    className="table-action"
                    onClick={exportBanhKem}
                    disabled={banhKemData.length === 0 || loadingState.banhKem}
                  >
                    <FiDownload />
                    Xuất Excel
                  </button>
                </div>
              </div>
              <div className="table-content">
                <div className="inventory-summary">
                  <p>Tổng số đơn bánh kem: {banhKemData.length}</p>
                  <p>Tổng số lượng: {formatNumber(totalBanhKemProducts)}</p>
                  <p>Tổng doanh thu: {formatCurrency(totalBanhKemRevenue)}</p>
                </div>
                {banhKemData.length > 0 ? (
                  <div className="modern-table">
                    <table>
                      <thead>
                        <tr>
                          <th>ID Điều Chuyển</th>
                          <th>Sale Order</th>
                          <th>Mã SP</th>
                          <th>Tên Sản Phẩm</th>
                          <th>Số Lượng</th>
                          <th>Đơn Giá</th>
                          <th>Thành Tiền</th>
                          <th>Ngày Thực Hiện</th>
                          <th>Nhân Viên</th>
                          <th>Cửa Hàng</th>
                        </tr>
                      </thead>
                      <tbody>
                        {banhKemData
                          .sort(
                            (a, b) =>
                              new Date(b.NgayThucHien) -
                              new Date(a.NgayThucHien)
                          )
                          .slice(0, 20)
                          .map((item, index) => {
                            const storeName =
                              stores.find(
                                (store) => store.MaCuaHang === item.MaCH
                              )?.TenCuaHang || item.MaCuaHang;
                            const localDate = new Date(
                              new Date(item.NgayThucHien).getTime() -
                                7 * 60 * 60 * 1000
                            );
                            return (
                              <tr key={`${item.IDDieuChuyen}-${index}`}>
                                <td>
                                  <span className="product-code">
                                    {item.IDDieuChuyen}
                                  </span>
                                </td>
                                <td>{item.Sale_Order}</td>
                                <td>
                                  <span className="product-code">
                                    {item.MaSP}
                                  </span>
                                </td>
                                <td className="product-name">{item.TenSP}</td>
                                <td className="quantity">
                                  {formatNumber(item.SoLuong)}
                                </td>
                                <td className="price">
                                  {formatCurrency(item.DonGia)}
                                </td>
                                <td className="amount">
                                  {formatCurrency(item.SoLuong * item.DonGia)}
                                </td>
                                <td>
                                  {localDate.toLocaleDateString("vi-VN")}{" "}
                                  {localDate.toLocaleTimeString("vi-VN", {
                                    hour: "2-digit",
                                    minute: "2-digit",
                                  })}
                                </td>
                                <td>
                                  <span className="product-code">
                                    {item.NhanVien}
                                  </span>
                                </td>
                                <td className="store-name">{storeName}</td>
                              </tr>
                            );
                          })}
                      </tbody>
                    </table>
                  </div>
                ) : (
                  <div className="empty-state">
                    <div className="empty-icon">
                      <FiPackage />
                    </div>
                    <h4>Chưa có dữ liệu bánh kem</h4>
                    <p>
                      Không có dữ liệu bán bánh kem trong khoảng thời gian đã
                      chọn
                    </p>
                  </div>
                )}
                {banhKemData.length > 20 && (
                  <p className="more-data">
                    ... và {banhKemData.length - 20} đơn bánh kem khác (xem
                    trong file Excel)
                  </p>
                )}
              </div>
            </div>

            {/* Bánh Kem Đặt Data Table */}
            <div className="table-card">
              <div className="table-header">
                <div className="table-title">
                  <FiPackage />
                  <h3>Chi Tiết Bánh Kem Đặt {loadingState.wholesale && <FiRefreshCw className="spinning" style={{ marginLeft: 8, fontSize: 14 }} />}</h3>
                </div>
                <div className="table-actions">
                  <button
                    className="table-action"
                    onClick={exportWholesale}
                    disabled={wholesaleData.length === 0 || loadingState.wholesale}
                  >
                    <FiDownload />
                    Xuất Excel
                  </button>
                </div>
              </div>
              <div className="table-content">
                <div className="inventory-summary">
                  <p>Tổng số đơn bánh kem đặt: {wholesaleData.length}</p>
                  <p>Tổng số lượng: {formatNumber(totalWholesaleProducts)}</p>
                  <p>Tổng doanh thu: {formatCurrency(totalWholesaleRevenue)}</p>
                </div>
                {wholesaleData.length > 0 ? (
                  <div className="modern-table">
                    <table>
                      <thead>
                        <tr>
                          <th>Mã PB</th>
                          <th>ID CH</th>
                          <th>Mã SP</th>
                          <th>Đơn Giá</th>
                          <th>Số Lượng</th>
                          <th>Thành Tiền</th>
                          <th>Ngày Đặt</th>
                          <th>Ngày Nhận</th>
                          <th>Người Đặt</th>
                          <th>Người Xác Nhận</th>
                        </tr>
                      </thead>
                      <tbody>
                        {wholesaleData
                          .sort(
                            (a, b) => new Date(b.NgayDat) - new Date(a.NgayDat)
                          )
                          .slice(0, 20)
                          .map((item, index) => {
                            const thanhTien = item.SoLuong * item.DonGia;
                            const ngayDat = new Date(item.NgayDat);
                            const ngayNhan = new Date(item.NgayNhan);
                            return (
                              <tr key={`${item.MaPB}-${index}`}>
                                <td>
                                  <span className="product-code">
                                    {item.MaPB}
                                  </span>
                                </td>
                                <td>{item.IDCH}</td>
                                <td>
                                  <span className="product-code">
                                    {item.MaSP}
                                  </span>
                                </td>
                                <td className="price">
                                  {formatCurrency(item.DonGia)}
                                </td>
                                <td className="quantity">
                                  {formatNumber(item.SoLuong)}
                                </td>
                                <td className="amount">
                                  {formatCurrency(thanhTien)}
                                </td>
                                <td>
                                  {ngayDat.toLocaleDateString("vi-VN")}{" "}
                                  {ngayDat.toLocaleTimeString("vi-VN", {
                                    hour: "2-digit",
                                    minute: "2-digit",
                                  })}
                                </td>
                                <td>
                                  {ngayNhan.toLocaleDateString("vi-VN")}{" "}
                                  {ngayNhan.toLocaleTimeString("vi-VN", {
                                    hour: "2-digit",
                                    minute: "2-digit",
                                  })}
                                </td>
                                <td>
                                  <span className="product-code">
                                    {item.NguoiDat}
                                  </span>
                                </td>
                                <td>
                                  <span className="product-code">
                                    {item.NguoiXacNhan}
                                  </span>
                                </td>
                              </tr>
                            );
                          })}
                      </tbody>
                    </table>
                  </div>
                ) : (
                  <div className="empty-state">
                    <div className="empty-icon">
                      <FiPackage />
                    </div>
                    <h4>Chưa có dữ liệu bánh kem đặt</h4>
                    <p>
                      Không có dữ liệu bánh kem đặt trong khoảng thời gian đã
                      chọn
                    </p>
                  </div>
                )}
                {wholesaleData.length > 20 && (
                  <p className="more-data">
                    ... và {wholesaleData.length - 20} đơn bánh kem đặt khác
                    (xem trong file Excel)
                  </p>
                )}
              </div>
            </div>

            <div className="table-card">
              <div className="table-header">
                <div className="table-title">
                  <FiPackage />
                  <h3>Thống Kê Sản Phẩm {loadingState.thongKe && <FiRefreshCw className="spinning" style={{ marginLeft: 8, fontSize: 14 }} />}</h3>
                </div>
              </div>
              <div className="table-content">
                <div className="modern-table">
                  <table>
                    <thead>
                      <tr>
                        <th>Mã SP</th>
                        <th>Tên Sản Phẩm</th>
                        <th>Số Lượng</th>
                        <th>Giá Bán</th>
                        <th>Thành Tiền</th>
                      </tr>
                    </thead>
                    <tbody>
                      {data.THDoanhSoSP.slice(0, 15).map((item, index) => (
                        <tr key={index}>
                          <td>
                            <span className="product-code">{item.MaSP}</span>
                          </td>
                          <td className="product-name">{item.TenSP}</td>
                          <td className="quantity">
                            {formatNumber(item.SoLuong)}
                          </td>
                          <td className="price">
                            {formatCurrency(item.GiaBan)}
                          </td>
                          <td className="amount">
                            {formatCurrency(item.ThanhTien)}
                          </td>
                        </tr>
                      ))}
                    </tbody>
                  </table>
                </div>
              </div>
            </div>

            <div className="table-card">
              <div className="table-header">
                <div className="table-title">
                  <FiFileText />
                  <h3>Sổ Chi Tiết Nhập Hàng {loadingState.nhapHang && <FiRefreshCw className="spinning" style={{ marginLeft: 8, fontSize: 14 }} />}</h3>
                </div>
              </div>
              <div className="table-content">
                <div className="nhap-hang-summary">
                  <p>Tổng số phiếu nhập: {nhapHangData.length}</p>
                  <p>
                    Tổng số mặt hàng:{" "}
                    {nhapHangData.reduce(
                      (sum, phieu) => sum + phieu.Items.length,
                      0
                    )}
                  </p>
                </div>
                {nhapHangData.length > 0 ? (
                  <div className="nhap-hang-list">
                    {nhapHangData.slice(0, 10).map((phieu, index) => (
                      <div key={index} className="nhap-hang-phieu">
                        <div className="phieu-header">
                          <strong>Phiếu: {phieu.SoPN || phieu.Phieu}</strong> |
                          Ngày:{" "}
                          {new Date(phieu.NgayNhap).toLocaleDateString("vi-VN")}{" "}
                          | Cửa hàng: {phieu.MaCH}
                        </div>
                        <div className="modern-table">
                          <table>
                            <thead>
                              <tr>
                                <th>Mã SP</th>
                                <th>Tên SP</th>
                                <th>Số lượng</th>
                                <th>Đơn giá</th>
                                <th>Giảm giá(%)</th>
                                <th>Thành tiền</th>
                              </tr>
                            </thead>
                            <tbody>
                              {phieu.Items.map((item, itemIndex) => (
                                <tr key={itemIndex}>
                                  <td>
                                    <span className="product-code">
                                      {item.MaSP}
                                    </span>
                                  </td>
                                  <td className="product-name">{item.TenSP}</td>
                                  <td className="quantity">
                                    {formatNumber(item.SoLuong)}
                                  </td>
                                  <td className="price">
                                    {formatCurrency(item.GiaMua)}
                                  </td>
                                  <td className="discount">
                                    {item.GiamPT || 0}%
                                  </td>
                                  <td className="amount">
                                    {formatCurrency(item.ThanhTien)}
                                  </td>
                                </tr>
                              ))}
                            </tbody>
                          </table>
                        </div>
                      </div>
                    ))}
                    {nhapHangData.length > 10 && (
                      <p className="more-data">
                        ... và {nhapHangData.length - 10} phiếu khác (xem trong
                        file Excel)
                      </p>
                    )}
                  </div>
                ) : (
                  <p>Không có dữ liệu nhập hàng trong khoảng thời gian này.</p>
                )}
              </div>
            </div>

            <div className="table-card">
              <div className="table-header">
                <div className="table-title">
                  <FiCreditCard />
                  <h3>Dữ Liệu Ca Làm {loadingState.ketCa && <FiRefreshCw className="spinning" style={{ marginLeft: 8, fontSize: 14 }} />}</h3>
                </div>
                <div className="table-actions">
                  <button
                    className="table-action"
                    onClick={exportKetCa}
                    disabled={shiftData.length === 0 || loadingState.ketCa}
                  >
                    <FiDownload />
                    Xuất Excel
                  </button>
                </div>
              </div>
              <div className="table-content">
                <div className="inventory-summary">
                  <p>Tổng số ngày: {shiftData.length}</p>
                  <p>Tổng số ca: {totalShifts}</p>
                  <p>Tổng doanh thu ca: {formatCurrency(totalShiftRevenue)}</p>
                </div>
                {shiftData.length > 0 ? (
                  <div className="modern-table">
                    <table>
                      <thead>
                        <tr>
                          <th>Ngày</th>
                          <th>Ca</th>
                          <th>Nhân Viên</th>
                          <th>Thời Gian Kết Ca</th>
                          <th>Tổng Tiền</th>
                          <th>Giảm Giá</th>
                          <th>Tiền Sau Giảm Giá</th>
                        </tr>
                      </thead>
                      <tbody>
                        {shiftData.slice(0, 15).map((dayData, dayIndex) =>
                          dayData.Ca.map((shift, shiftIndex) => (
                            <tr key={`${dayIndex}-${shiftIndex}`}>
                              <td>
                                {new Date(dayData.Ngay).toLocaleDateString(
                                  "vi-VN"
                                )}
                              </td>
                              <td>
                                <span className="product-code">
                                  Ca {shift.CaID}
                                </span>
                              </td>
                              <td className="product-name">
                                {shift.TenNhanVien}
                              </td>
                              <td>{shift.ThoiGianKetCa}</td>
                              <td className="amount">
                                {formatCurrency(
                                  shift.TongTien + (shift.DSBanhKem || 0)
                                )}
                              </td>
                              <td className="discount">
                                {formatCurrency(shift.GiamGia)}
                              </td>
                              <td className="net-amount">
                                {formatCurrency(
                                  shift.TienSauGiamGia + (shift.DSBanhKem || 0)
                                )}
                              </td>
                            </tr>
                          ))
                        )}
                      </tbody>
                    </table>
                  </div>
                ) : (
                  <div className="empty-state">
                    <div className="empty-icon">
                      <FiCreditCard />
                    </div>
                    <h4>Chưa có dữ liệu ca làm</h4>
                    <p>
                      Không có dữ liệu ca làm trong khoảng thời gian đã chọn
                    </p>
                  </div>
                )}
                {(() => {
                  const totalDisplayedShifts = shiftData
                    .slice(0, 15)
                    .reduce((sum, day) => sum + day.Ca.length, 0);
                  return (
                    totalShifts > totalDisplayedShifts && (
                      <p className="more-data">
                        ... và {totalShifts - totalDisplayedShifts} ca làm khác
                        (xem trong file Excel)
                      </p>
                    )
                  );
                })()}
              </div>
            </div>

            {/* Inventory Table */}
            <div className="table-card">
              <div className="table-header">
                <div className="table-title">
                  <FiPackage />
                  <h3>Thống Kê Tồn Kho</h3>
                </div>
                <div className="table-actions">
                  <button
                    className="table-action"
                    onClick={async () => {
                      if (khoHangData.length === 0) {
                        await fetchKhoHang();
                      }
                    }}
                    disabled={loadingState.khoHang}
                  >
                    <FiRefreshCw className={loadingState.khoHang ? "spinning" : ""} />
                    {loadingState.khoHang ? "Đang tải..." : "Tải dữ liệu tồn kho"}
                  </button>
                  <button
                    className="table-action"
                    onClick={exportTonKho}
                    disabled={loadingState.khoHang}
                  >
                    <FiDownload />
                    Xuất Excel
                  </button>
                </div>
              </div>
              <div className="table-content">
                <div className="inventory-summary">
                  {khoHangData.length > 0 ? (
                    <>
                      <p>Tổng số mặt hàng: {khoHangData.length}</p>
                      <p>Tổng số lượng tồn: {formatNumber(totalInventoryItems)}</p>
                      <p>Tổng giá trị tồn: {formatCurrency(totalInventoryValue)}</p>
                    </>
                  ) : (
                    <p>Nhấn "Tải dữ liệu tồn kho" hoặc "Xuất Excel" để lấy dữ liệu</p>
                  )}
                </div>
                {khoHangData.length > 0 ? (
                  <div className="modern-table">
                    <table>
                      <thead>
                        <tr>
                          <th>STT</th>
                          <th>Mã SP</th>
                          <th>Tên Sản Phẩm</th>
                          <th>ĐVT</th>
                          <th>Tồn đầu kỳ</th>
                          <th>GT đầu kỳ</th>
                          <th>Tổng nhập</th>
                          <th>GT nhập</th>
                          {!showSplitXuat ? (
                            <th
                              style={{
                                cursor: "pointer",
                                textDecoration: "underline",
                              }}
                              onClick={() => setShowSplitXuat(true)}
                              title="Nhấn để xem chi tiết xuất/hủy"
                            >
                              Tổng xuất
                            </th>
                          ) : (
                            <>
                              <th
                                style={{
                                  cursor: "pointer",
                                  textDecoration: "underline",
                                }}
                                onClick={() => setShowSplitXuat(false)}
                                title="Nhấn để gộp lại"
                              >
                                Tổng xuất
                              </th>
                              <th
                                style={{
                                  cursor: "pointer",
                                  textDecoration: "underline",
                                }}
                                onClick={() => setShowSplitXuat(false)}
                                title="Nhấn để gộp lại"
                              >
                                Tổng hủy
                              </th>
                            </>
                          )}
                          {!showSplitGiaTriXuat ? (
                            <th
                              style={{
                                cursor: "pointer",
                                textDecoration: "underline",
                              }}
                              onClick={() => setShowSplitGiaTriXuat(true)}
                              title="Nhấn để xem chi tiết GT xuất/hủy"
                            >
                              GT xuất
                            </th>
                          ) : (
                            <>
                              <th
                                style={{
                                  cursor: "pointer",
                                  textDecoration: "underline",
                                }}
                                onClick={() => setShowSplitGiaTriXuat(false)}
                                title="Nhấn để gộp lại"
                              >
                                GT xuất
                              </th>
                              <th
                                style={{
                                  cursor: "pointer",
                                  textDecoration: "underline",
                                }}
                                onClick={() => setShowSplitGiaTriXuat(false)}
                                title="Nhấn để gộp lại"
                              >
                                GT hủy
                              </th>
                            </>
                          )}
                          <th>Tồn cuối kỳ</th>
                          <th>GT cuối kỳ</th>
                        </tr>
                      </thead>
                      <tbody>
                        {khoHangData.slice(0, 20).map((item, index) => {
                          const transferInfo = getTransferDataForProduct(item);
                          return (
                            <tr key={item.MaNL || index}>
                              <td>{item.STT || index + 1}</td>
                              <td>
                                <span className="product-code">
                                  {item.MaNL}
                                </span>
                              </td>
                              <td className="product-name">{item.Ten_SP}</td>
                              <td>{item.DVT_SP}</td>
                              <td className="quantity">
                                {formatNumber(item.TonDK || item.TonDauKy || 0)}
                              </td>
                              <td className="amount">
                                {formatCurrency(
                                  item.GiaTriDK || item.GiaTriDauKy || 0
                                )}
                              </td>
                              <td className="quantity">
                                {formatNumber(
                                  (item.TongNhap || 0) +
                                    transferInfo.nhapDieuChuyen
                                )}
                              </td>
                              <td className="amount">
                                {formatCurrency(
                                  (item.GiaTriNhap || 0) +
                                    transferInfo.giaTriNhapDieuChuyen
                                )}
                              </td>
                              {!showSplitXuat ? (
                                <td className="quantity">
                                  {formatNumber(
                                    (item.TongXuat || 0) + (item.TongHuy || 0)
                                  )}
                                </td>
                              ) : (
                                <>
                                  <td className="quantity">
                                    {formatNumber(item.TongXuat || 0)}
                                  </td>
                                  <td className="quantity">
                                    {formatNumber(item.TongHuy || 0)}
                                  </td>
                                </>
                              )}
                              {!showSplitGiaTriXuat ? (
                                <td className="amount">
                                  {formatCurrency(
                                    (item.GiaTriXuat || 0) +
                                      (item.GiaTriHuy || 0)
                                  )}
                                </td>
                              ) : (
                                <>
                                  <td className="amount">
                                    {formatCurrency(item.GiaTriXuat || 0)}
                                  </td>
                                  <td className="amount">
                                    {formatCurrency(item.GiaTriHuy || 0)}
                                  </td>
                                </>
                              )}
                              <td className="quantity inventory-current">
                                {formatNumber(item.TonCK || 0)}
                              </td>
                              <td className="amount inventory-value">
                                {formatCurrency(item.GiaTriCK || 0)}
                              </td>
                            </tr>
                          );
                        })}
                      </tbody>
                    </table>
                  </div>
                ) : (
                  <div className="empty-state">
                    <div className="empty-icon">
                      <FiPackage />
                    </div>
                    <h4>Chưa có dữ liệu tồn kho</h4>
                    <p>
                      Vui lòng chọn cửa hàng và thời gian để xem thống kê tồn
                      kho
                    </p>
                  </div>
                )}
                {khoHangData.length > 20 && (
                  <p className="more-data">
                    ... và {khoHangData.length - 20} mặt hàng khác (xem trong
                    file Excel)
                  </p>
                )}
              </div>
            </div>

            {/* Inventory Check Table */}
            <div className="table-card">
              <div className="table-header">
                <div className="table-title">
                  <FiClipboard />
                  <h3>Phiếu Kiểm Kê {loadingState.kiemKe && <FiRefreshCw className="spinning" style={{ marginLeft: 8, fontSize: 14 }} />}</h3>
                </div>
                <div className="table-actions">
                  <button
                    className="table-action"
                    onClick={exportKiemKe}
                    disabled={kiemKeData.length === 0 || loadingState.kiemKe}
                  >
                    <FiDownload />
                    Xuất Excel chi tiết
                  </button>
                </div>
              </div>
              <div className="table-content">
                <div className="inventory-summary">
                  <p>Tổng số phiếu: {kiemKeData.length}</p>
                  <p>
                    Khoảng thời gian: {filters.tuNgay} - {filters.denNgay}
                  </p>
                  <p>Sắp xếp: Theo ngày tăng dần</p>
                </div>
                {kiemKeData.length > 0 ? (
                  <div className="modern-table">
                    <table>
                      <thead>
                        <tr>
                          <th>ID Phiếu</th>
                          <th>Ngày Kiểm Kê</th>
                          <th>Người Kiểm Kê</th>
                          <th>Ghi Chú</th>
                        </tr>
                      </thead>
                      <tbody>
                        {[...kiemKeData]
                          .sort(
                            (a, b) =>
                              new Date(a.NgayKiemKe) - new Date(b.NgayKiemKe)
                          )
                          .slice(0, 15)
                          .map((phieu, index) => (
                            <tr key={phieu.IDPhieu}>
                              <td>
                                <span className="product-code">
                                  #{phieu.IDPhieu}
                                </span>
                              </td>
                              <td>
                                {new Date(phieu.NgayKiemKe).toLocaleDateString(
                                  "vi-VN"
                                )}
                              </td>
                              <td className="product-name">{phieu.TenUser}</td>
                              <td>{phieu.GhiChu || "-"}</td>
                            </tr>
                          ))}
                      </tbody>
                    </table>
                  </div>
                ) : (
                  <div className="empty-state">
                    <div className="empty-icon">
                      <FiClipboard />
                    </div>
                    <h4>Chưa có phiếu kiểm kê</h4>
                    <p>
                      Không có phiếu kiểm kê nào trong khoảng thời gian đã chọn
                    </p>
                  </div>
                )}
                {kiemKeData.length > 15 && (
                  <p className="more-data">
                    ... và {kiemKeData.length - 15} phiếu kiểm kê khác (xem
                    trong file Excel)
                  </p>
                )}
              </div>
            </div>
          </div>
        </>
    </div>
  );
};

export default Dashboard;
