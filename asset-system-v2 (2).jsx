import React, { useState, useEffect, useCallback, useRef } from "react";
import * as XLSX from "xlsx";
import { Laptop, Smartphone, Printer, Car, Wifi, Package, CheckCircle2, Users, FileSpreadsheet, Trash2, Search, ChevronDown, AlertCircle, Lock, Image } from "lucide-react";

// ---------- Config ----------
const ADMIN_PASSWORD = "1997";
const STORAGE_KEY = "darbco_asset_records_v2";

// ---------- Static reference data ----------
const EMPLOYEES = [
  { name: "Monther Fadel", dept: "Executive Management", position: "CEO" },
  { name: "Amjad Khlail", dept: "Executive Management", position: "CTO" },
  { name: "Shireen Khaleel", dept: "Operations", position: "Operation Manager" },
  { name: "Alaedeen Mustafa", dept: "Sales", position: "Sales Manager" },
  { name: "Ibrahem Alwahdani", dept: "Manufacturing", position: "Site Engineer" },
  { name: "Yousef Absi", dept: "Manufacturing", position: "Design Engineer" },
  { name: "Mohammad Alqasi", dept: "Manufacturing", position: "QC Engineer" },
  { name: "Alaedeen Alrababah", dept: "Manufacturing", position: "Production Engineer" },
  { name: "Ahmad Aladdasi", dept: "Sales", position: "Sales Engineer" },
  { name: "Nadia Fadel", dept: "Operations", position: "HR Manager" },
  { name: "Farah Odeh", dept: "Marketing", position: "Marketing Officer" },
  { name: "Gana Saliby", dept: "Operations", position: "After Sales Engineer" },
  { name: "Omar Adel", dept: "Operations", position: "Procurement Engineer" },
  { name: "Dima Shaltaf", dept: "Sales", position: "Sales Engineer" },
  { name: "Rami Abdallah", dept: "Sales", position: "Sales Engineer" },
  { name: "Mohammad Waleed Hammad", dept: "Finance", position: "Accountant" },
];

const DEPARTMENTS = [
  "Executive Management", "Marketing", "Operations", "Sales",
  "Production", "Procurement", "Finance", "Manufacturing",
];

const ASSET_TYPES = [
  { value: "Computer/Laptop", label: "Computer / Laptop", labelAr: "أجهزة كمبيوتر / لابتوب", icon: Laptop },
  { value: "Phone/Mobile", label: "Phone / Mobile Device", labelAr: "هواتف وأجهزة موبايل", icon: Smartphone },
  { value: "Office Equipment", label: "Office Equipment (Printer, Monitor)", labelAr: "معدات مكتبية (طابعات، شاشات)", icon: Printer },
  { value: "Vehicle", label: "Vehicle", labelAr: "سيارات / مركبات", icon: Car },
  { value: "Sim Card", label: "Sim Card", labelAr: "Sim Card", icon: Wifi },
  { value: "Other", label: "Other", labelAr: "أخرى", icon: Package },
];

const ASSET_ICON_MAP = ASSET_TYPES.reduce((acc, t) => { acc[t.value] = t.icon; return acc; }, {});

const EMPTY_FORM = {
  employeeName: "", employeeNameOther: "",
  department: "", departmentOther: "",
  position: "", assetType: "", assetTypeOther: "",
  assetModel: "", serialNumber: "", handoverDate: "", notes: "",
};

// ---------- Helpers ----------
function genId() { return "rec_" + Date.now() + "_" + Math.random().toString(36).slice(2, 8); }
function todayISO() { return new Date().toISOString().slice(0, 10); }

function loadFromStorage() {
  try {
    const raw = localStorage.getItem(STORAGE_KEY);
    if (!raw) return [];
    const parsed = JSON.parse(raw);
    return Array.isArray(parsed) ? parsed : [];
  } catch { return []; }
}

function saveToStorage(records) {
  try {
    localStorage.setItem(STORAGE_KEY, JSON.stringify(records));
    return true;
  } catch { return false; }
}

// Convert file to base64 for storage
function fileToBase64(file) {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = () => resolve(reader.result);
    reader.onerror = reject;
    reader.readAsDataURL(file);
  });
}

function useIsNarrow() {
  const [narrow, setNarrow] = useState(typeof window !== "undefined" ? window.innerWidth < 640 : false);
  useEffect(() => {
    const onResize = () => setNarrow(window.innerWidth < 640);
    window.addEventListener("resize", onResize);
    return () => window.removeEventListener("resize", onResize);
  }, []);
  return narrow;
}

// ---------- Main App ----------
export default function AssetManagementSystem() {
  const [view, setView] = useState("form");
  const [adminUnlocked, setAdminUnlocked] = useState(false);
  const [passwordInput, setPasswordInput] = useState("");
  const [passwordError, setPasswordError] = useState(false);
  const [form, setForm] = useState(EMPTY_FORM);
  const [imageFile, setImageFile] = useState(null);
  const [imagePreview, setImagePreview] = useState(null);
  const [records, setRecords] = useState(() => loadFromStorage());
  const [submitting, setSubmitting] = useState(false);
  const [submitted, setSubmitted] = useState(false);
  const [errors, setErrors] = useState({});
  const [search, setSearch] = useState("");
  const [filterDept, setFilterDept] = useState("");
  const [filterType, setFilterType] = useState("");
  const [confirmDelete, setConfirmDelete] = useState(null);
  const isNarrow = useIsNarrow();

  const handleEmployeeSelect = (value) => {
    if (value === "__other__") {
      setForm((f) => ({ ...f, employeeName: "__other__", employeeNameOther: "", department: "", position: "" }));
      return;
    }
    const emp = EMPLOYEES.find((e) => e.name === value);
    setForm((f) => ({ ...f, employeeName: value, department: emp?.dept || "", position: emp?.position || "" }));
  };

  const handleImageChange = (e) => {
    const file = e.target.files[0];
    if (!file) return;
    setImageFile(file);
    setImagePreview(URL.createObjectURL(file));
    if (errors.image) setErrors((prev) => { const n = { ...prev }; delete n.image; return n; });
  };

  const validate = () => {
    const e = {};
    const finalName = form.employeeName === "__other__" ? form.employeeNameOther.trim() : form.employeeName;
    if (!finalName) e.employeeName = "Required";
    if (form.employeeName === "__other__" && !form.employeeNameOther.trim()) e.employeeNameOther = "Required";
    if (!form.department) e.department = "Required";
    if (form.department === "__other__" && !form.departmentOther.trim()) e.departmentOther = "Required";
    const finalAsset = form.assetType === "__other__" ? form.assetTypeOther.trim() : form.assetType;
    if (!finalAsset) e.assetType = "Required";
    if (form.assetType === "__other__" && !form.assetTypeOther.trim()) e.assetTypeOther = "Required";
    if (!form.assetModel.trim()) e.assetModel = "Required";
    if (!form.serialNumber.trim()) e.serialNumber = "Required";
    if (!form.handoverDate) e.handoverDate = "Required";
    if (!imageFile) e.image = "Required — يجب رفع صورة ملصق الجهاز";
    setErrors(e);
    return Object.keys(e).length === 0;
  };

  const handleSubmit = async (ev) => {
    ev.preventDefault();
    if (!validate()) return;
    setSubmitting(true);

    const finalEmployeeName = form.employeeName === "__other__" ? form.employeeNameOther.trim() : form.employeeName;
    const finalDepartment = form.department === "__other__" ? form.departmentOther.trim() : form.department;
    const finalAssetType = form.assetType === "__other__" ? form.assetTypeOther.trim() : form.assetType;

    let imageData = null;
    try { imageData = await fileToBase64(imageFile); } catch { /* ignore */ }

    const newRecord = {
      id: genId(),
      employeeName: finalEmployeeName,
      department: finalDepartment,
      position: form.position || "",
      assetType: finalAssetType,
      assetModel: form.assetModel.trim(),
      serialNumber: form.serialNumber.trim(),
      handoverDate: form.handoverDate,
      notes: form.notes.trim(),
      imageData,
      imageName: imageFile?.name || "",
      submittedAt: new Date().toISOString(),
    };

    const next = [...records, newRecord];
    const ok = saveToStorage(next);
    if (ok) {
      setRecords(next);
      setSubmitted(true);
      setForm(EMPTY_FORM);
      setImageFile(null);
      setImagePreview(null);
      setTimeout(() => setSubmitted(false), 4000);
    }
    setSubmitting(false);
  };

  const handlePasswordSubmit = (e) => {
    e.preventDefault();
    if (passwordInput === ADMIN_PASSWORD) {
      setAdminUnlocked(true);
      setPasswordError(false);
    } else {
      setPasswordError(true);
      setPasswordInput("");
    }
  };

  const deleteRecord = (id) => {
    const next = records.filter((r) => r.id !== id);
    saveToStorage(next);
    setRecords(next);
    setConfirmDelete(null);
  };

  const exportToExcel = () => {
    const rows = filteredRecords.map((r, idx) => ({
      "Asset ID": "DARBCO-" + String(idx + 1).padStart(4, "0"),
      "Employee Name": r.employeeName,
      Department: r.department,
      Position: r.position,
      "Asset Type": r.assetType,
      "Asset Model": r.assetModel,
      "Serial Number": r.serialNumber,
      "Handover Date": r.handoverDate,
      Notes: r.notes,
      "Image Attached": r.imageData ? "Yes" : "No",
      "Submitted At": new Date(r.submittedAt).toLocaleString(),
    }));
    const ws = XLSX.utils.json_to_sheet(rows);
    ws["!cols"] = [14,22,20,20,20,22,18,14,30,14,20].map(w => ({ wch: w }));
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "Asset Register");
    XLSX.writeFile(wb, `DARBCO_Asset_Register_${todayISO()}.xlsx`);
  };

  const filteredRecords = records.filter((r) => {
    const matchesSearch = !search ||
      r.employeeName.toLowerCase().includes(search.toLowerCase()) ||
      r.assetModel.toLowerCase().includes(search.toLowerCase()) ||
      r.serialNumber.toLowerCase().includes(search.toLowerCase());
    return matchesSearch && (!filterDept || r.department === filterDept) && (!filterType || r.assetType === filterType);
  });

  const stats = {
    total: records.length,
    byType: ASSET_TYPES.map((t) => ({ ...t, count: records.filter((r) => r.assetType === t.value).length })),
    departments: new Set(records.map((r) => r.department)).size,
  };

  const handleAdminTabClick = () => {
    if (!adminUnlocked) { setView("admin"); }
    else { setView("admin"); }
  };

  return (
    <div style={styles.page}>
      <style>{fontImport}</style>
      <div style={styles.shell}>
        <Header view={view} setView={setView} isNarrow={isNarrow} adminUnlocked={adminUnlocked} />
        {view === "form" ? (
          <EmployeeForm
            form={form} setForm={setForm} errors={errors}
            submitting={submitting} submitted={submitted}
            handleSubmit={handleSubmit} handleEmployeeSelect={handleEmployeeSelect}
            imageFile={imageFile} imagePreview={imagePreview} handleImageChange={handleImageChange}
            isNarrow={isNarrow}
          />
        ) : !adminUnlocked ? (
          <AdminLogin
            passwordInput={passwordInput} setPasswordInput={setPasswordInput}
            passwordError={passwordError} handlePasswordSubmit={handlePasswordSubmit}
          />
        ) : (
          <AdminDashboard
            records={records} filteredRecords={filteredRecords} stats={stats}
            search={search} setSearch={setSearch}
            filterDept={filterDept} setFilterDept={setFilterDept}
            filterType={filterType} setFilterType={setFilterType}
            exportToExcel={exportToExcel}
            confirmDelete={confirmDelete} setConfirmDelete={setConfirmDelete}
            deleteRecord={deleteRecord} isNarrow={isNarrow}
            onLogout={() => { setAdminUnlocked(false); setPasswordInput(""); setView("form"); }}
          />
        )}
      </div>
    </div>
  );
}

// ---------- Header ----------
function Header({ view, setView, isNarrow, adminUnlocked }) {
  return (
    <div style={{ ...styles.header, ...(isNarrow ? { flexDirection: "column", alignItems: "stretch" } : {}) }}>
      <div style={styles.brandRow}>
        <div style={styles.brandTag}><span style={styles.brandTagCode}>DR-01</span></div>
        <div>
          <div style={styles.brandName}>DARBCO Robotics</div>
          <div style={styles.brandSub}>Asset Custody System · نظام إدارة العهد</div>
        </div>
      </div>
      <div style={{ ...styles.navTabs, ...(isNarrow ? { width: "100%" } : {}) }}>
        <button onClick={() => setView("form")}
          style={{ ...styles.navTab, ...(isNarrow ? { flex: 1 } : {}), ...(view === "form" ? styles.navTabActive : {}) }}>
          Submit Asset · تسجيل عهدة
        </button>
        <button onClick={() => setView("admin")}
          style={{ ...styles.navTab, ...(isNarrow ? { flex: 1 } : {}), ...(view === "admin" ? styles.navTabActive : {}) }}>
          <Lock size={13} style={{ marginLeft: 4, opacity: adminUnlocked ? 0.4 : 1 }} />
          HR Dashboard · لوحة الإدارة
        </button>
      </div>
    </div>
  );
}

// ---------- Admin Login ----------
function AdminLogin({ passwordInput, setPasswordInput, passwordError, handlePasswordSubmit }) {
  return (
    <div style={styles.formWrap}>
      <div style={{ ...styles.formCard, maxWidth: 400 }}>
        <div style={{ textAlign: "center", marginBottom: 24 }}>
          <div style={{ width: 52, height: 52, borderRadius: 12, background: COLORS.primary + "15", display: "flex", alignItems: "center", justifyContent: "center", margin: "0 auto 14px" }}>
            <Lock size={24} color={COLORS.primary} />
          </div>
          <h2 style={{ ...styles.formTitle, fontSize: 19 }}>HR Dashboard Access</h2>
          <p style={{ ...styles.formSubtitle, marginTop: 6 }}>لوحة الإدارة خاصة بـ HR فقط — أدخلي كلمة السر للمتابعة</p>
        </div>
        {passwordError && (
          <div style={styles.errorBanner}>
            <AlertCircle size={16} style={{ flexShrink: 0 }} />
            <span>كلمة السر غير صحيحة · Incorrect password</span>
          </div>
        )}
        <form onSubmit={handlePasswordSubmit} style={{ display: "flex", flexDirection: "column", gap: 14 }}>
          <input
            type="password" value={passwordInput}
            onChange={(e) => setPasswordInput(e.target.value)}
            placeholder="كلمة السر · Password"
            autoFocus
            style={{ ...styles.input, textAlign: "center", letterSpacing: 4, fontSize: 18 }}
          />
          <button type="submit" style={styles.submitBtn}>دخول · Enter</button>
        </form>
      </div>
    </div>
  );
}

// ---------- Employee Form ----------
function EmployeeForm({ form, setForm, errors, submitting, submitted, handleSubmit, handleEmployeeSelect, imageFile, imagePreview, handleImageChange, isNarrow }) {
  const set = (key) => (e) => setForm((f) => ({ ...f, [key]: e.target.value }));
  const twoCol = isNarrow ? { ...styles.twoCol, gridTemplateColumns: "1fr" } : styles.twoCol;
  const assetGrid = isNarrow ? { ...styles.assetTypeGrid, gridTemplateColumns: "repeat(2, 1fr)" } : styles.assetTypeGrid;
  const cardStyle = isNarrow ? { ...styles.formCard, padding: "24px 18px 28px" } : styles.formCard;
  const fileInputRef = useRef();

  return (
    <div style={styles.formWrap}>
      <div style={cardStyle}>
        <div style={styles.formCardHeader}>
          <h2 style={styles.formTitle}>Register your assigned asset</h2>
          <p style={styles.formSubtitle}>سجّل أي جهاز، هاتف، أو عهدة تم تسليمها لك. يستخدم قسم الموارد البشرية هذه البيانات لإدارة سجل العهد في الشركة.</p>
        </div>
        {submitted && (
          <div style={styles.successBanner}>
            <CheckCircle2 size={18} style={{ flexShrink: 0 }} />
            <span>Recorded successfully · تم تسجيل العهدة بنجاح</span>
          </div>
        )}
        <form onSubmit={handleSubmit} style={styles.form}>
          {/* Employee Name */}
          <div style={styles.fieldGroup}>
            <Field label="Employee Name" labelAr="اسم الموظف" error={errors.employeeName} required>
              <Select value={form.employeeName} onChange={(e) => handleEmployeeSelect(e.target.value)} error={errors.employeeName}>
                <option value="">Select your name…</option>
                {EMPLOYEES.map((emp) => <option key={emp.name} value={emp.name}>{emp.name}</option>)}
                <option value="__other__">Other (type manually) · أخرى</option>
              </Select>
            </Field>
            {form.employeeName === "__other__" && (
              <Field label="Enter your name" labelAr="اكتب اسمك" error={errors.employeeNameOther} required>
                <Input value={form.employeeNameOther} onChange={set("employeeNameOther")} placeholder="Full name" />
              </Field>
            )}
          </div>

          {/* Dept + Position */}
          <div style={twoCol}>
            <Field label="Department" labelAr="القسم" hint={form.employeeName && form.employeeName !== "__other__" ? "Auto-filled" : null}>
              {form.employeeName && form.employeeName !== "__other__" ? (
                <ReadOnlyChip value={form.department} />
              ) : (
                <Select value={form.department} onChange={set("department")} error={errors.department}>
                  <option value="">Select department…</option>
                  {DEPARTMENTS.map((d) => <option key={d} value={d}>{d}</option>)}
                  <option value="__other__">Other · أخرى</option>
                </Select>
              )}
            </Field>
            <Field label="Position" labelAr="المسمى الوظيفي">
              {form.employeeName && form.employeeName !== "__other__" ? (
                <ReadOnlyChip value={form.position || "—"} />
              ) : (
                <Input value={form.position} onChange={set("position")} placeholder="Job title" />
              )}
            </Field>
          </div>
          {form.department === "__other__" && (
            <Field label="Enter department" labelAr="اكتب القسم" error={errors.departmentOther} required>
              <Input value={form.departmentOther} onChange={set("departmentOther")} placeholder="Department name" />
            </Field>
          )}

          {/* Asset Type */}
          <Field label="Asset Type" labelAr="نوع العهدة" error={errors.assetType} required>
            <div style={assetGrid}>
              {ASSET_TYPES.map((t) => {
                const Icon = t.icon;
                const active = form.assetType === t.value;
                return (
                  <button type="button" key={t.value}
                    onClick={() => setForm((f) => ({ ...f, assetType: t.value }))}
                    style={{ ...styles.assetTypeCard, ...(active ? styles.assetTypeCardActive : {}) }}>
                    <Icon size={20} strokeWidth={1.75} />
                    <span style={styles.assetTypeLabel}>{t.label}</span>
                    <span style={styles.assetTypeLabelAr}>{t.labelAr}</span>
                  </button>
                );
              })}
            </div>
          </Field>
          {form.assetType === "Other" && (
            <Field label="Specify asset type" labelAr="حدد نوع العهدة" error={errors.assetTypeOther} required>
              <Input value={form.assetTypeOther} onChange={set("assetTypeOther")} placeholder="e.g. Tablet, Tool kit" />
            </Field>
          )}

          {/* Model + Serial */}
          <div style={twoCol}>
            <Field label="Asset Model / Name" labelAr="اسم / موديل الجهاز" error={errors.assetModel} required>
              <Input value={form.assetModel} onChange={set("assetModel")} placeholder="e.g. Lenovo ThinkPad T14" />
            </Field>
            <Field label="Serial Number" labelAr="الرقم التسلسلي" error={errors.serialNumber} required>
              <Input value={form.serialNumber} onChange={set("serialNumber")} placeholder="As printed on device label" />
            </Field>
          </div>

          {/* Image Upload — REQUIRED */}
          <Field label="Device Label Photo" labelAr="صورة ملصق الجهاز (الرقم التسلسلي)" error={errors.image} required>
            <input ref={fileInputRef} type="file" accept="image/*" onChange={handleImageChange} style={{ display: "none" }} />
            <div
              onClick={() => fileInputRef.current.click()}
              style={{
                ...styles.imageUploadBox,
                ...(errors.image ? { borderColor: COLORS.danger } : {}),
                ...(imagePreview ? { borderStyle: "solid", borderColor: COLORS.primary } : {}),
              }}
            >
              {imagePreview ? (
                <div style={styles.imagePreviewWrap}>
                  <img src={imagePreview} alt="preview" style={styles.imagePreviewImg} />
                  <div style={styles.imagePreviewName}>{imageFile?.name}</div>
                  <div style={{ fontSize: 11.5, color: COLORS.primary, fontWeight: 600, marginTop: 4 }}>اضغط لتغيير الصورة · Click to change</div>
                </div>
              ) : (
                <div style={styles.imageUploadPlaceholder}>
                  <Image size={28} strokeWidth={1.5} color={COLORS.inkSoft} />
                  <span style={{ fontSize: 13.5, fontWeight: 600, color: COLORS.ink, marginTop: 8 }}>ارفع صورة ملصق الجهاز</span>
                  <span style={{ fontSize: 12, color: COLORS.inkSoft }}>Upload photo of device label · اضغط هنا</span>
                </div>
              )}
            </div>
          </Field>

          {/* Date + Notes */}
          <div style={twoCol}>
            <Field label="Handover Date" labelAr="تاريخ التسليم" error={errors.handoverDate} required>
              <Input type="date" value={form.handoverDate} onChange={set("handoverDate")} />
            </Field>
            <Field label="Notes (optional)" labelAr="ملاحظات (اختياري)">
              <Input value={form.notes} onChange={set("notes")} placeholder="Condition, accessories, etc." />
            </Field>
          </div>

          <button type="submit" disabled={submitting} style={styles.submitBtn}>
            {submitting ? "Saving…" : "Submit · إرسال"}
          </button>
        </form>
      </div>
    </div>
  );
}

function ReadOnlyChip({ value }) { return <div style={styles.readOnlyChip}>{value}</div>; }

function Field({ label, labelAr, children, error, required, hint }) {
  return (
    <div style={styles.field}>
      <label style={styles.label}>
        <span>{label} {required && <span style={styles.required}>*</span>}</span>
        <span style={styles.labelAr}>{labelAr}</span>
        {hint && <span style={styles.hint}>{hint}</span>}
      </label>
      {children}
      {error && <span style={styles.errorText}>{error}</span>}
    </div>
  );
}

function Input(props) { return <input {...props} style={{ ...styles.input, ...(props.style || {}) }} />; }

function Select({ children, error, ...props }) {
  return (
    <div style={styles.selectWrap}>
      <select {...props} style={{ ...styles.input, ...styles.select, ...(error ? styles.inputError : {}) }}>{children}</select>
      <ChevronDown size={16} style={styles.selectChevron} />
    </div>
  );
}

// ---------- Admin Dashboard ----------
function AdminDashboard({ records, filteredRecords, stats, search, setSearch, filterDept, setFilterDept, filterType, setFilterType, exportToExcel, confirmDelete, setConfirmDelete, deleteRecord, isNarrow, onLogout }) {
  const [viewingImage, setViewingImage] = useState(null);
  const statsRowStyle = isNarrow ? { ...styles.statsRow, gridTemplateColumns: "1fr 1fr" } : styles.statsRow;

  return (
    <div style={styles.adminWrap}>
      {/* Image lightbox */}
      {viewingImage && (
        <div onClick={() => setViewingImage(null)} style={styles.lightboxOverlay}>
          <div onClick={(e) => e.stopPropagation()} style={styles.lightboxBox}>
            <img src={viewingImage} alt="asset" style={styles.lightboxImg} />
            <button onClick={() => setViewingImage(null)} style={styles.lightboxClose}>✕ إغلاق</button>
          </div>
        </div>
      )}

      <div style={{ display: "flex", justifyContent: "flex-end", marginBottom: -6 }}>
        <button onClick={onLogout} style={styles.logoutBtn}>
          <Lock size={13} /> تسجيل خروج
        </button>
      </div>

      <div style={statsRowStyle}>
        <StatCard icon={Package} label="Total Assets" labelAr="إجمالي العهد" value={stats.total} accent="#1E5F5F" />
        <StatCard icon={Users} label="Departments" labelAr="الأقسام المغطاة" value={stats.departments} accent="#C9802B" />
        <StatCard icon={Laptop} label="Computers/Laptops" labelAr="كمبيوتر / لابتوب" value={stats.byType.find((t) => t.value === "Computer/Laptop")?.count || 0} accent="#3D5A80" />
        <StatCard icon={Smartphone} label="Phones/Mobiles" labelAr="هواتف / موبايل" value={stats.byType.find((t) => t.value === "Phone/Mobile")?.count || 0} accent="#8B5A8C" />
      </div>

      <div style={styles.toolbar}>
        <div style={styles.searchWrap}>
          <Search size={16} style={styles.searchIcon} />
          <input value={search} onChange={(e) => setSearch(e.target.value)}
            placeholder="Search by name, model, or serial number…" style={styles.searchInput} />
        </div>
        <select value={filterDept} onChange={(e) => setFilterDept(e.target.value)} style={styles.filterSelect}>
          <option value="">All Departments</option>
          {DEPARTMENTS.map((d) => <option key={d} value={d}>{d}</option>)}
        </select>
        <select value={filterType} onChange={(e) => setFilterType(e.target.value)} style={styles.filterSelect}>
          <option value="">All Asset Types</option>
          {ASSET_TYPES.map((t) => <option key={t.value} value={t.value}>{t.label}</option>)}
        </select>
        <button onClick={exportToExcel} disabled={filteredRecords.length === 0} style={styles.exportBtn}>
          <FileSpreadsheet size={16} /> Export to Excel
        </button>
      </div>

      <div style={styles.tableCard}>
        {filteredRecords.length === 0 ? (
          <div style={styles.emptyState}>
            <Package size={32} strokeWidth={1.5} style={{ opacity: 0.35, marginBottom: 8 }} />
            <div>{records.length === 0 ? "No assets registered yet." : "No records match your filters."}</div>
          </div>
        ) : (
          <div style={styles.tableScroll}>
            <table style={styles.table}>
              <thead>
                <tr>
                  {["Employee", "Department", "Position", "Asset Type", "Model", "Serial No.", "Photo", "Handover Date", "Notes", ""].map((h) => (
                    <th key={h} style={styles.th}>{h}</th>
                  ))}
                </tr>
              </thead>
              <tbody>
                {filteredRecords.slice().reverse().map((r) => {
                  const Icon = ASSET_ICON_MAP[r.assetType] || Package;
                  return (
                    <tr key={r.id} style={styles.tr}>
                      <td style={styles.td}><span style={styles.empName}>{r.employeeName}</span></td>
                      <td style={styles.td}>{r.department}</td>
                      <td style={styles.td}>{r.position}</td>
                      <td style={styles.td}>
                        <span style={styles.typeChip}><Icon size={13} strokeWidth={1.75} />{r.assetType}</span>
                      </td>
                      <td style={styles.td}>{r.assetModel}</td>
                      <td style={styles.td}><code style={styles.serialCode}>{r.serialNumber}</code></td>
                      <td style={styles.td}>
                        {r.imageData ? (
                          <img src={r.imageData} alt="label" onClick={() => setViewingImage(r.imageData)}
                            style={styles.thumbImg} title="Click to enlarge" />
                        ) : <span style={{ color: COLORS.inkSoft, fontSize: 12 }}>—</span>}
                      </td>
                      <td style={styles.td}>{r.handoverDate}</td>
                      <td style={{ ...styles.td, ...styles.notesCell }}>{r.notes || "—"}</td>
                      <td style={styles.td}>
                        {confirmDelete === r.id ? (
                          <div style={styles.confirmRow}>
                            <button onClick={() => deleteRecord(r.id)} style={styles.confirmBtn}>Confirm</button>
                            <button onClick={() => setConfirmDelete(null)} style={styles.cancelBtn}>Cancel</button>
                          </div>
                        ) : (
                          <button onClick={() => setConfirmDelete(r.id)} style={styles.deleteBtn} title="Delete">
                            <Trash2 size={14} />
                          </button>
                        )}
                      </td>
                    </tr>
                  );
                })}
              </tbody>
            </table>
          </div>
        )}
      </div>
      <div style={styles.footnote}>Showing {filteredRecords.length} of {records.length} record{records.length === 1 ? "" : "s"}</div>
    </div>
  );
}

function StatCard({ icon: Icon, label, labelAr, value, accent }) {
  return (
    <div style={styles.statCard}>
      <div style={{ ...styles.statIconWrap, background: accent + "1A", color: accent }}><Icon size={18} strokeWidth={1.75} /></div>
      <div>
        <div style={styles.statValue}>{value}</div>
        <div style={styles.statLabel}>{label}</div>
        <div style={styles.statLabelAr}>{labelAr}</div>
      </div>
    </div>
  );
}

// ---------- Styles ----------
const fontImport = `@import url('https://fonts.googleapis.com/css2?family=Manrope:wght@400;500;600;700;800&family=IBM+Plex+Mono:wght@500&display=swap');`;

const COLORS = {
  bg: "#F6F4EF", paper: "#FFFFFF", ink: "#1C2526", inkSoft: "#5B6566",
  rule: "#E2DED3", primary: "#1E5F5F", primaryDark: "#15494A",
  accent: "#C9802B", danger: "#B3432B",
  successBg: "#E8F2EC", successText: "#1E5F3F", errorBg: "#FBEAE5",
};

const styles = {
  page: { fontFamily: "'Manrope', system-ui, sans-serif", background: COLORS.bg, minHeight: "100vh", color: COLORS.ink, padding: 0, boxSizing: "border-box" },
  shell: { maxWidth: 1100, margin: "0 auto", padding: "20px 16px 48px" },
  header: { display: "flex", flexWrap: "wrap", alignItems: "center", justifyContent: "space-between", gap: 16, marginBottom: 28, paddingBottom: 18, borderBottom: `1px solid ${COLORS.rule}` },
  brandRow: { display: "flex", alignItems: "center", gap: 14 },
  brandTag: { width: 46, height: 46, borderRadius: 8, background: COLORS.primary, display: "flex", alignItems: "center", justifyContent: "center", boxShadow: "inset 0 0 0 3px rgba(255,255,255,0.18)", flexShrink: 0 },
  brandTagCode: { fontFamily: "'IBM Plex Mono', monospace", fontSize: 10, fontWeight: 600, color: "#fff", letterSpacing: 0.5 },
  brandName: { fontSize: 18, fontWeight: 800, letterSpacing: -0.2 },
  brandSub: { fontSize: 12.5, color: COLORS.inkSoft, marginTop: 2 },
  navTabs: { display: "flex", gap: 6, background: "#fff", padding: 4, borderRadius: 10, border: `1px solid ${COLORS.rule}` },
  navTab: { border: "none", background: "transparent", padding: "9px 16px", borderRadius: 7, fontSize: 13.5, fontWeight: 600, color: COLORS.inkSoft, cursor: "pointer", fontFamily: "inherit", display: "flex", alignItems: "center", gap: 4 },
  navTabActive: { background: COLORS.primary, color: "#fff" },
  formWrap: { display: "flex", justifyContent: "center" },
  formCard: { background: COLORS.paper, borderRadius: 14, border: `1px solid ${COLORS.rule}`, padding: "32px 36px 36px", width: "100%", maxWidth: 680, boxShadow: "0 1px 2px rgba(28,37,38,0.04)" },
  formCardHeader: { marginBottom: 22 },
  formTitle: { fontSize: 21, fontWeight: 800, margin: 0, letterSpacing: -0.2 },
  formSubtitle: { fontSize: 13.5, color: COLORS.inkSoft, marginTop: 6, lineHeight: 1.6 },
  form: { display: "flex", flexDirection: "column", gap: 18 },
  fieldGroup: { display: "flex", flexDirection: "column", gap: 18 },
  twoCol: { display: "grid", gridTemplateColumns: "1fr 1fr", gap: 16 },
  field: { display: "flex", flexDirection: "column", gap: 6 },
  label: { display: "flex", flexWrap: "wrap", alignItems: "baseline", gap: 8, fontSize: 13, fontWeight: 700, color: COLORS.ink },
  labelAr: { fontSize: 12.5, fontWeight: 500, color: COLORS.inkSoft },
  required: { color: COLORS.danger },
  hint: { fontSize: 11, fontWeight: 500, color: COLORS.primary, background: COLORS.primary + "14", padding: "1px 7px", borderRadius: 5 },
  input: { fontFamily: "inherit", fontSize: 14, padding: "10px 12px", borderRadius: 8, border: `1.5px solid ${COLORS.rule}`, outline: "none", background: "#fff", color: COLORS.ink, width: "100%", boxSizing: "border-box" },
  inputError: { borderColor: COLORS.danger },
  selectWrap: { position: "relative" },
  select: { appearance: "none", paddingRight: 34, cursor: "pointer" },
  selectChevron: { position: "absolute", right: 12, top: "50%", transform: "translateY(-50%)", color: COLORS.inkSoft, pointerEvents: "none" },
  readOnlyChip: { fontSize: 14, padding: "10px 12px", borderRadius: 8, background: "#F1EFE8", border: `1.5px dashed ${COLORS.rule}`, color: COLORS.inkSoft, fontWeight: 600 },
  errorText: { fontSize: 11.5, color: COLORS.danger, fontWeight: 600 },
  assetTypeGrid: { display: "grid", gridTemplateColumns: "repeat(3, 1fr)", gap: 10 },
  assetTypeCard: { display: "flex", flexDirection: "column", alignItems: "flex-start", gap: 6, padding: "12px 12px", borderRadius: 9, border: `1.5px solid ${COLORS.rule}`, background: "#fff", cursor: "pointer", fontFamily: "inherit", color: COLORS.ink, textAlign: "left", transition: "all 0.12s ease" },
  assetTypeCardActive: { borderColor: COLORS.primary, background: COLORS.primary + "0D", boxShadow: `0 0 0 1px ${COLORS.primary}` },
  assetTypeLabel: { fontSize: 12.5, fontWeight: 700, lineHeight: 1.25 },
  assetTypeLabelAr: { fontSize: 11, color: COLORS.inkSoft, fontWeight: 500 },
  imageUploadBox: { border: `2px dashed ${COLORS.rule}`, borderRadius: 10, padding: "20px 16px", cursor: "pointer", background: "#FAFAF8", transition: "border-color 0.15s ease", textAlign: "center" },
  imageUploadPlaceholder: { display: "flex", flexDirection: "column", alignItems: "center", gap: 4 },
  imagePreviewWrap: { display: "flex", flexDirection: "column", alignItems: "center", gap: 6 },
  imagePreviewImg: { maxHeight: 120, maxWidth: "100%", borderRadius: 8, objectFit: "contain", border: `1px solid ${COLORS.rule}` },
  imagePreviewName: { fontSize: 12, color: COLORS.inkSoft },
  submitBtn: { marginTop: 6, background: COLORS.primary, color: "#fff", border: "none", borderRadius: 9, padding: "13px 20px", fontSize: 14.5, fontWeight: 700, cursor: "pointer", fontFamily: "inherit" },
  successBanner: { display: "flex", alignItems: "center", gap: 8, background: COLORS.successBg, color: COLORS.successText, padding: "10px 14px", borderRadius: 8, fontSize: 13.5, fontWeight: 600, marginBottom: 18 },
  errorBanner: { display: "flex", alignItems: "center", gap: 8, background: COLORS.errorBg, color: COLORS.danger, padding: "10px 14px", borderRadius: 8, fontSize: 13.5, fontWeight: 600, marginBottom: 18 },
  adminWrap: { display: "flex", flexDirection: "column", gap: 18 },
  statsRow: { display: "grid", gridTemplateColumns: "repeat(4, 1fr)", gap: 12 },
  statCard: { background: COLORS.paper, border: `1px solid ${COLORS.rule}`, borderRadius: 12, padding: "14px 16px", display: "flex", alignItems: "center", gap: 12 },
  statIconWrap: { width: 38, height: 38, borderRadius: 9, display: "flex", alignItems: "center", justifyContent: "center", flexShrink: 0 },
  statValue: { fontSize: 20, fontWeight: 800, lineHeight: 1.1 },
  statLabel: { fontSize: 11.5, color: COLORS.inkSoft, fontWeight: 600, marginTop: 2 },
  statLabelAr: { fontSize: 11, color: COLORS.inkSoft, opacity: 0.8 },
  toolbar: { display: "flex", flexWrap: "wrap", gap: 10, alignItems: "center" },
  searchWrap: { position: "relative", flex: "1 1 240px", minWidth: 200 },
  searchIcon: { position: "absolute", left: 11, top: "50%", transform: "translateY(-50%)", color: COLORS.inkSoft },
  searchInput: { width: "100%", boxSizing: "border-box", fontFamily: "inherit", fontSize: 13.5, padding: "10px 12px 10px 34px", borderRadius: 8, border: `1.5px solid ${COLORS.rule}`, outline: "none", background: "#fff" },
  filterSelect: { fontFamily: "inherit", fontSize: 13.5, padding: "10px 12px", borderRadius: 8, border: `1.5px solid ${COLORS.rule}`, outline: "none", background: "#fff", color: COLORS.ink, cursor: "pointer" },
  exportBtn: { display: "flex", alignItems: "center", gap: 7, background: COLORS.accent, color: "#fff", border: "none", borderRadius: 8, padding: "10px 16px", fontSize: 13.5, fontWeight: 700, cursor: "pointer", fontFamily: "inherit", whiteSpace: "nowrap" },
  logoutBtn: { display: "flex", alignItems: "center", gap: 6, background: "transparent", border: `1px solid ${COLORS.rule}`, color: COLORS.inkSoft, borderRadius: 7, padding: "7px 12px", fontSize: 12.5, fontWeight: 600, cursor: "pointer", fontFamily: "inherit" },
  tableCard: { background: COLORS.paper, border: `1px solid ${COLORS.rule}`, borderRadius: 12, overflow: "hidden" },
  tableScroll: { overflowX: "auto" },
  table: { width: "100%", borderCollapse: "collapse", fontSize: 13 },
  th: { textAlign: "left", padding: "11px 14px", fontSize: 11.5, fontWeight: 700, color: COLORS.inkSoft, textTransform: "uppercase", letterSpacing: 0.4, borderBottom: `1.5px solid ${COLORS.rule}`, whiteSpace: "nowrap" },
  tr: { borderBottom: `1px solid ${COLORS.rule}` },
  td: { padding: "11px 14px", verticalAlign: "middle", color: COLORS.ink },
  empName: { fontWeight: 700 },
  typeChip: { display: "inline-flex", alignItems: "center", gap: 5, fontSize: 11.5, fontWeight: 600, background: "#F1EFE8", padding: "3px 8px", borderRadius: 6, whiteSpace: "nowrap" },
  serialCode: { fontFamily: "'IBM Plex Mono', monospace", fontSize: 12, background: "#F1EFE8", padding: "2px 6px", borderRadius: 5 },
  thumbImg: { width: 48, height: 48, objectFit: "cover", borderRadius: 6, cursor: "pointer", border: `1px solid ${COLORS.rule}` },
  notesCell: { maxWidth: 180, color: COLORS.inkSoft, fontSize: 12.5 },
  deleteBtn: { border: "none", background: "transparent", color: COLORS.inkSoft, cursor: "pointer", padding: 6, borderRadius: 6, display: "flex" },
  confirmRow: { display: "flex", gap: 6 },
  confirmBtn: { border: "none", background: COLORS.danger, color: "#fff", fontSize: 11.5, fontWeight: 700, padding: "5px 9px", borderRadius: 6, cursor: "pointer", fontFamily: "inherit" },
  cancelBtn: { border: `1px solid ${COLORS.rule}`, background: "#fff", color: COLORS.inkSoft, fontSize: 11.5, fontWeight: 600, padding: "5px 9px", borderRadius: 6, cursor: "pointer", fontFamily: "inherit" },
  emptyState: { padding: "48px 20px", textAlign: "center", color: COLORS.inkSoft, fontSize: 13.5, display: "flex", flexDirection: "column", alignItems: "center" },
  footnote: { fontSize: 12, color: COLORS.inkSoft, textAlign: "right" },
  lightboxOverlay: { position: "fixed", inset: 0, background: "rgba(0,0,0,0.75)", display: "flex", alignItems: "center", justifyContent: "center", zIndex: 9999, padding: 20 },
  lightboxBox: { background: "#fff", borderRadius: 12, padding: 16, maxWidth: "90vw", maxHeight: "90vh", display: "flex", flexDirection: "column", alignItems: "center", gap: 12 },
  lightboxImg: { maxWidth: "80vw", maxHeight: "70vh", objectFit: "contain", borderRadius: 8 },
  lightboxClose: { border: "none", background: COLORS.ink, color: "#fff", borderRadius: 7, padding: "8px 20px", cursor: "pointer", fontFamily: "inherit", fontSize: 14, fontWeight: 600 },
};
