import * as React from "react";
//import "./advanced.scss";
import { spfi } from "@pnp/sp";
import { SPFx } from "@pnp/sp/presets/all";
import { useEffect, useState } from "react";
import {
  PeoplePicker,
  PrincipalType,
} from "@pnp/spfx-controls-react/lib/PeoplePicker";
import { IPeoplePickerContext } from "@pnp/spfx-controls-react/lib/PeoplePicker";
import logo from "../assets/sona-comstarlogo.png";
import "bootstrap/dist/css/bootstrap.min.css";
interface IVendor {
  Id: number;
  VendorCode: string;
  VendorName: string;
}

const EditAdvanceForm = ({ context, formData, onClose }: any) => {
  const sp = spfi().using(SPFx(context));

  // =========================
  // STATES
  // =========================
  const [pickerKey, setPickerKey] = useState<number>(0);
  const [previousAdvances, setPreviousAdvances] = useState<any[]>([]);
  const [vendors, setVendors] = useState<IVendor[]>([]);
  const [employee, setEmployee] = useState<any>({});
  const [attachments, setAttachments] = useState<any[]>([]);
  const [selectedFiles, setSelectedFiles] = useState<File[]>([]);
  const [selectedUser, setSelectedUser] = useState<any[]>([]);

  const [selectedVendorId, setSelectedVendorId] = useState<number | null>(null);
  const [selectedVendorName, setSelectedVendorName] = useState("");

  const [poNumber, setPoNumber] = useState("");
  const [poDate, setPoDate] = useState("");
  const [poTerms, setPoTerms] = useState("");
  const [poAmount, setPoAmount] = useState("");
  const [advanceAmount, setAdvanceAmount] = useState("");
  const [paidAmount, setPaidAmount] = useState("");
  const [expectedDate, setExpectedDate] = useState("");
  const [vendorName, setVendorName] = useState("");
  const [glCode, setGlCode] = useState("390111001");
  const [costCenter, setCostCenter] = useState("");
  const [remarks, setRemarks] = useState("");
  const [projectDesc, setProjectDesc] = useState("");

  const [approvalMatrix, setApprovalMatrix] = useState<any[]>([]);
  const [workflowHistory, setWorkflowHistory] = useState<any[]>([]);

  const peoplePickerContext: IPeoplePickerContext = {
    absoluteUrl: context.pageContext.web.absoluteUrl,
    msGraphClientFactory: context.msGraphClientFactory,
    spHttpClient: context.spHttpClient,
  };

  // =========================
  // LOAD DATA
  // =========================
  const getPreviousAdvances = async (vendorId: number) => {
    try {
      debugger;
      console.log("Fetching for Vendor:", vendorId);

      const data = await sp.web.lists
        .getByTitle("OpexAdvance")
        .items.select(
          "PONumber",
          "RequestAdvanceAmount",
          "Created",
          "VoucherDate",

          "PaidAmount",
          "Status",
          "VendorCode/Id",
        )
        .expand("VendorCode")
        .filter(`VendorCode/Id eq ${vendorId} and Status eq 'Paid'`)
        .orderBy("Created", false)();

      console.log("DATA:", data);

      void setPreviousAdvances(data);
    } catch (error) {
      console.error("Error fetching previous advances:", error);
      void setPreviousAdvances([]);
    }
  };
  const handleNumberChange = (value: string, setter: any) => {
    // Allow only numbers and decimal (max one dot)
    const regex = /^\d*\.?\d*$/;

    if (regex.test(value)) {
      void setter(value);
    }
  };
  const getVendors = async () => {
    const data = await sp.web.lists
      .getByTitle("VendorMaster")
      .items.select("Id", "VendorCode", "VendorName")();
    void setVendors(data);
  };

  const getLoggedInUser = async () => {
    try {
      const currentUser = await sp.web.currentUser();
      const email = currentUser.Email;

      const user = await sp.web.lists
        .getByTitle("EmployeeMaster")
        .items.select(
          "EmployeeCode",
          "EmployeeName",
          "Division",
          "Location",
          "EmployeeEmail",
          "ReportingManager/Title",
          "HOD/Title",
          "ContactNo",
          "EmployeeStatus",
          "CostCenter",
        )
        .expand("ReportingManager", "HOD")
        .filter(`EmployeeEmail eq '${email}'`)
        .top(1)();

      if (user.length > 0) {
        setEmployee(user[0]);
      }
    } catch (error) {
      console.log("Error fetching user:", error);
    }
  };
  const buildApprovalFlow = async () => {
    const flow: any[] = [];

    // 🔹 RM
    if (employee.ReportingManager?.Id) {
      flow.push({
        Id: employee.ReportingManager.Id,
        Name: employee.ReportingManager.Title,
        Role: "RM",
        Level: 1,
        Status: "Pending",
      });
    }

    // 🔹 HOD
    if (employee.HOD?.Id) {
      flow.push({
        Id: employee.HOD.Id,
        Name: employee.HOD.Title,
        Role: "HOD",
        Level: 2,
        Status: "Pending",
      });
    }

    // 🔹 Matrix approvers
    const matrixData = await sp.web.lists
      .getByTitle("CapexApprovalMatrix")
      .items.select("Role/RoleName,Approver/Id,Approver/Title,Level/Level")
      .expand("Approver,Role,Level")
      .filter("Status eq 'Active'")
      .orderBy("Level", true)();

    const matrixApprovers = matrixData.map((item: any, index: number) => ({
      Id: item.Approver?.Id,
      Name: item.Approver?.Title,
      Role: item.Role?.RoleName,
      Level: flow.length + index + 1,
      Status: "Pending",
    }));

    const finalFlow = [...flow, ...matrixApprovers];

    // 🔥 first approver active
    if (finalFlow.length > 0) {
      finalFlow[0].Status = "In Progress";
    }

    return finalFlow;
  };
  const handledraft = async () => {
    try {

      debugger;


      const flow = await buildApprovalFlow();
      const currentApprover = flow.length > 0 ? flow[0].Id : null;
      let ensuredUserId: number | null = null;
      const existingFlow = formData.ApprovalMatrix
        ? JSON.parse(formData.ApprovalMatrix)
        : [];

      // 🔥 preserve history
      const history = formData.WorkFlowHistory
        ? JSON.parse(formData.WorkFlowHistory)
        : [];

      history.push({
        CurrentApprover: employee.EmployeeName,
        ActionTaken: "Edited",
        Comment: remarks,
        Date: new Date().toISOString(),
      });

      await sp.web.lists
        .getByTitle("OpexAdvance")
        .items.getById(formData.ID)
        .update({
          Title: formData.CapexID,
          CapexID: formData.CapexID,

          EmployeeCode: employee.EmployeeCode,
          EmployeeName: employee.EmployeeName,

          VendorCodeId: selectedVendorId,
          VendorName: selectedVendorName,

          PONumber: poNumber,
          PODate: poDate ? new Date(poDate) : null,

          POAmtGST: poAmount,
          RequestAdvanceAmount: advanceAmount,
          PaidAmount: paidAmount,

          ExpectedDateofSettlement: expectedDate
            ? new Date(expectedDate)
            : null,

          PICNameId: ensuredUserId,

          GL: glCode,
          CostCenter: costCenter,
          Remarks: remarks,
          Purpose: projectDesc,

          Status: "Draft", // ✅ important change
          ApprovalMatrix: JSON.stringify(flow),

          CurrentApproverId: currentApprover, // 🔥 not started

          WorkFlowHistory: JSON.stringify(history),
        });

      if (selectedFiles.length > 0) {
        await uploadFiles();
      }

      alert("Draft saved successfully ✅");
      window.location.href =
        "https://isriglobal.sharepoint.com/sites/SonaFinance/SitePages/OpexAdvancedForm.aspx?page=User";
      void handleExit();
      // window.location.reload();

    } catch (error) {
      console.error("ERROR:", error);
      alert("Error while saving ❌");
    }
  };

  const handleDeleteExistingFile = async (file: any) => {
    try {
      if (!window.confirm(`Delete ${file.Name}?`)) return;

      await sp.web.getFileByServerRelativePath(file.ServerRelativeUrl).delete();

      // update UI
      setAttachments((prev) =>
        prev.filter((f) => f.ServerRelativeUrl !== file.ServerRelativeUrl),
      );

      alert("File deleted ✅");
    } catch (error) {
      console.error("Delete error:", error);
    }
  };

  const getAttachments = async (capexId: string) => {
    try {
      const safe = capexId.replace(/\//g, "_");
      const path = `/sites/SonaFinance/Opex/${safe}`;

      const files = await sp.web.getFolderByServerRelativePath(path).files();

      void setAttachments(files);
    } catch {
      void setAttachments([]);
    }
  };

  // =========================
  // UPLOAD FILES
  // =========================

  // =========================
  // VALIDATION
  // =========================
  const validateForm = () => {
    const errors: string[] = [];

    if (!selectedVendorId) {
      errors.push("Please select the Vendor code");
    }

    if (!poNumber) {
      errors.push("Please update PO Number");
    }

    if (!poDate) {
      errors.push("Please update PO date");
    }
    if (!glCode) {
      errors.push("Please update GL Code");
    }
    if (!poTerms) {
      errors.push("Please update PO Terms");
    }

    if (!poAmount) {
      errors.push("Please update PO Amount");
    }

    if (!advanceAmount || Number(advanceAmount) <= 0) {
      errors.push("Please update Advance Amount");
    }

    if (!paidAmount || Number(paidAmount) <= 0) {
      errors.push("Please update Paid Amount");
    }

    //glCode
    // 🔥 NEW VALIDATION
    if (
      advanceAmount &&
      paidAmount &&
      Number(paidAmount) > Number(advanceAmount)
    ) {
      errors.push("Paid Amount cannot be greater than Advance Amount");
    }

    if (!expectedDate) {
      errors.push("Please update Settlement Date");
    }

    if (expectedDate) {
      const today = new Date().setHours(0, 0, 0, 0);
      const selected = new Date(expectedDate).setHours(0, 0, 0, 0);

      if (selected < today) {
        errors.push("Settlement date cannot be a past date");
      }
    }
    if (!selectedUser || selectedUser.length === 0) {
      errors.push("Please select PIC Name");
    }

    if (!projectDesc) {
      errors.push("Please enter Purpose");
    }

    if (
      (!attachments || attachments.length === 0) &&
      (!selectedFiles || selectedFiles.length === 0)
    ) {
      errors.push("Please upload at least one attachment");
    }

    return errors;
  };

  const handleExit = () => {
    if (onClose) {
      onClose();
    } else {
      window.location.reload();
    }
  };
  const ensureFolder = async (folderPath: string) => {
    try {
      await sp.web.getFolderByServerRelativePath(folderPath)();
    } catch {
      // create folder if not exists
      const parentPath = folderPath.substring(0, folderPath.lastIndexOf("/"));
      const folderName = folderPath.substring(folderPath.lastIndexOf("/") + 1);

      await sp.web
        .getFolderByServerRelativePath(parentPath)
        .folders.addUsingPath(folderName);
    }
  };
  const uploadFiles = async () => {
    try {
      if (!formData?.CapexID || selectedFiles.length === 0) return;

      const safe = formData.CapexID.replace(/\//g, "_");
      const folderPath = `/sites/SonaFinance/Opex/${safe}`;

      // ✅ Ensure folder exists
      await ensureFolder(folderPath);

      for (const file of selectedFiles) {
        await sp.web
          .getFolderByServerRelativePath(folderPath)
          .files.addUsingPath(file.name, file, { Overwrite: true });
      }

      setSelectedFiles([]);
      await getAttachments(formData.CapexID);
    } catch (error) {
      console.error("Upload error:", error);
      alert("File upload failed ❌");
    }
  };

  // =========================
  // UPDATE
  // =========================
  const handleSubmit = async () => {
    try {
      const errors = validateForm();
      if (errors.length > 0) {
        alert(errors.join("\n"));
        return;
      }

      const ensuredUser = await sp.web.ensureUser(
        selectedUser[0]?.secondaryText,
      );
      // 🔥 preserve matrix (no reset)
      const existingFlow = formData.ApprovalMatrix
        ? JSON.parse(formData.ApprovalMatrix)
        : [];

      // 🔥 preserve history
      const history = formData.WorkFlowHistory
        ? JSON.parse(formData.WorkFlowHistory)
        : [];

      // 🔥 add edit entry
      history.push({
        CurrentApprover: employee.EmployeeName,
        ActionTaken: "Edited",
        Comment: remarks,
        Date: new Date().toISOString(),
      });
      const currentApproverId = formData.CurrentApproverId || null;

      await sp.web.lists
        .getByTitle("OpexAdvance")
        .items.getById(formData.ID)
        .update({
          Title: formData.CapexID,
          CapexID: formData.CapexID,

          EmployeeCode: employee.EmployeeCode,
          EmployeeName: employee.EmployeeName,
          Division: employee.Division,
          Location: employee.Location,
          Email: employee.EmployeeEmail,
          RM: employee.ReportingManager?.Title,
          HOD: employee.HOD?.Title,
          ContactNo: employee.ContactNo,
          EmployeeStatus: employee.EmployeeStatus,

          VendorCodeId: selectedVendorId,
          VendorName: selectedVendorName,

          PONumber: poNumber,
          PODate: poDate ? new Date(poDate) : null,
          POAdvanceTerms: poTerms,

          POAmtGST: poAmount,
          RequestAdvanceAmount: advanceAmount,
          PaidAmount: paidAmount,

          ExpectedDateofSettlement: expectedDate
            ? new Date(expectedDate)
            : null,

          PICNameId: ensuredUser.Id,

          GL: glCode,
          CostCenter: costCenter,
          Remarks: remarks,
          Purpose: projectDesc,

          Status: "Pending for Approver",
          ApprovalMatrix: JSON.stringify(existingFlow),
          WorkFlowHistory: JSON.stringify(history),
          CurrentApproverId: currentApproverId,
        });

      if (selectedFiles.length > 0) {
        await uploadFiles();
      }

      alert("Updated successfully ✅");
      window.location.href =
        "https://isriglobal.sharepoint.com/sites/SonaFinance/SitePages/OpexAdvancedForm.aspx?page=User";
      void handleExit();
    } catch (e) {
      console.error(e);
      alert("Error ❌");
    }
  };

  // =========================
  // BIND DATA
  // =========================
  useEffect(() => {
    if (!formData) return;
    debugger;
    setPoNumber(formData.PONumber || "");
    setPoDate(formData.PODate?.split("T")[0] || "");
    setPoTerms(formData.POAdvanceTerms || "");
    setPoAmount(formData.POAmtGST || "");
    setAdvanceAmount(formData.RequestAdvanceAmount || "");
    setPaidAmount(formData.PaidAmount || "");
    setExpectedDate(formData.ExpectedDateofSettlement?.split("T")[0] || "");

    setVendorName(formData.VendorName || "");
    setSelectedVendorId(formData.VendorCodeId || null); // ✅ ADD THIS
    if (formData.VendorCodeId) {
      void getPreviousAdvances(formData.VendorCodeId);
    }
    setSelectedVendorName(formData.VendorName || ""); // ✅ ADD THIS

    setGlCode(formData.GL || "");
    setCostCenter(formData.CostCenter || "");
    setRemarks(formData.Remarks || "");
    setProjectDesc(formData.Purpose || "");

    // ✅ PIC FIX
    if (formData?.PICName?.Title) {
      setSelectedUser([
        {
          text: formData.PICName.Title,
          secondaryText: formData.PICName.EMail,
        },
      ]);
    }

    if (formData.CapexID) {
      void getAttachments(formData.CapexID);
    }
    // ✅ Approval Matrix
    if (formData?.ApprovalMatrix) {
      try {
        setApprovalMatrix(JSON.parse(formData.ApprovalMatrix));
      } catch (e) {
        console.error("ApprovalMatrix parse error", e);
      }
    }

    // ✅ Workflow History
    if (formData?.WorkFlowHistory) {
      try {
        setWorkflowHistory(JSON.parse(formData.WorkFlowHistory));
      } catch (e) {
        console.error("WorkFlowHistory parse error", e);
      }
    }
  }, [formData]);

  useEffect(() => {
    void getLoggedInUser();
    void getVendors();
    debugger;
    if (selectedVendorId) {
      void getPreviousAdvances(selectedVendorId);
    }
  }, []);


  // =========================
  // UI
  // =========================
  return (
    <div className="MainUplodForm" style={{ margin: "5px 0px" }}>
      <div className="row">
        <div className="col-md-12">
          <div className="Main-Boxpoup">
            {/* 🔹 Header */}
            <div className="bordered">
              <img src={logo} />
              <h1>Edit Advance Payment </h1>
            </div>
            {approvalMatrix.length === 0 ? (
              <p>No approval data</p>
            ) : (
              <div className="displayWF">
                <ul className="approval-flow">
                  {approvalMatrix.map((a, index) => (
                    <li
                      key={index}
                      className={`approval-step ${a.Status === "In Progress"
                        ? "active"
                        : a.Status === "Approved"
                          ? "approved"
                          : ""
                        }`}
                    >
                      {a.Role} - {a.Name}
                    </li>
                  ))}
                </ul>
              </div>
            )}
            <div className="borderedbox">
              <div className="heading1">
                <label>Requestor Information</label>
              </div>
              <div className="main-formcontainer">
                <div className="row mb-20">
                  <div className="col-md-4">
                    <label htmlFor="Employee Code" className="font">
                      Employee Code
                    </label>{" "}
                    : &nbsp;&nbsp;
                    <label className="fonttext"> {employee.EmployeeCode}</label>
                  </div>
                  <div className="col-md-4">
                    <label htmlFor="Employee Name" className="font">
                      Employee Name{" "}
                    </label>{" "}
                    : &nbsp;&nbsp;
                    <label className="fonttext"> {employee.EmployeeName}</label>
                  </div>
                  <div className="col-md-4">
                    <label htmlFor="Employee Email" className="font">
                      Employee Email{" "}
                    </label>{" "}
                    : &nbsp;&nbsp;
                    <label className="fonttext">
                      {" "}
                      {employee.EmployeeEmail}
                    </label>
                  </div>
                </div>
                <div className="row mb-20">
                  <div className="col-md-4">
                    <label htmlFor="Contact No" className="font">
                      Contact No
                    </label>{" "}
                    : &nbsp;&nbsp;
                    <label className="fonttext"> {employee.ContactNo}</label>
                  </div>
                  <div className="col-md-4">
                    <label htmlFor="Employee Status" className="font">
                      Employee Status
                    </label>{" "}
                    : &nbsp;&nbsp;
                    <label className="fonttext">
                      {" "}
                      {employee.EmployeeStatus}
                    </label>
                  </div>
                  <div className="col-md-4">
                    <label htmlFor="Division" className="font">
                      Division
                    </label>{" "}
                    : &nbsp;&nbsp;
                    <label className="fonttext"> {employee.Division}</label>
                  </div>
                </div>
                <div className="row mb-20">
                  <div className="col-md-4">
                    <label htmlFor="Location" className="font">
                      Location
                    </label>{" "}
                    : &nbsp;&nbsp;
                    <label className="fonttext"> {employee.Location}</label>
                  </div>
                  <div className="col-md-4">
                    <label htmlFor="RM" className="font">
                      RM
                    </label>{" "}
                    : &nbsp;&nbsp;
                    <label className="fonttext">
                      {" "}
                      {employee.ReportingManager?.Title}
                    </label>
                  </div>
                  <div className="col-md-4">
                    <label htmlFor="HOD" className="font">
                      HOD
                    </label>{" "}
                    : &nbsp;&nbsp;
                    <label className="fonttext"> {employee.HOD?.Title}</label>
                  </div>
                </div>
              </div>
              <div className="heading1" style={{ marginTop: "10px" }}>
                <label>Opex Details</label>
              </div>
              <div className="main-formcontainer">
                <div className="row mb-20">
                  <div className="col-md-4">
                    <label className="font">Vendor Code</label>
                    <select
                      value={selectedVendorId || ""}
                      onChange={(e) => {
                        const id = Number(e.target.value);
                        const vendor = vendors.find((v) => v.Id === id);

                        setSelectedVendorId(id);
                        setSelectedVendorName(vendor?.VendorName || "");

                        // 🔥 IMPORTANT: Fetch previous advances
                        if (id) {
                          void getPreviousAdvances(id);
                        }
                      }}
                      className="formtext-control"
                    >
                      <option value="">Select Vendor</option>
                      {vendors.map((v) => (
                        <option key={v.Id} value={v.Id}>
                          {v.VendorCode}
                        </option>
                      ))}
                    </select>

                    {/* <select
                      value={selectedVendorId || ""}
                      onChange={(e) => {
                        const id = Number(e.target.value);
                        const vendor = vendors.find((v) => v.Id === id);
                        setSelectedVendorId(id);
                        setSelectedVendorName(vendor?.VendorName || "");
                      }}
                      className="formtext-control"
                    >
                      <option value="">Select Vendor</option>
                      {vendors.map((v) => (
                        <option key={v.Id} value={v.Id}>
                          {v.VendorCode}
                        </option>
                      ))}
                    </select> */}
                  </div>
                  <div className="col-md-4">
                    <label className="font">Vendor Name</label>
                    <input
                      value={selectedVendorName || vendorName}
                      className="form-control readonly"
                    />
                  </div>
                  <div className="col-md-4">
                    <label className="font">PO Number</label>
                    <input
                      value={poNumber}
                      onChange={(e) => setPoNumber(e.target.value)}
                      className="form-control"
                    />
                  </div>
                </div>
                <div className="row mb-20">
                  <div className="col-md-4">
                    <label className="font">PO Date</label>
                    <input
                      type="date"
                      value={poDate}
                      onChange={(e) => setPoDate(e.target.value)}
                      className="form-control"
                    />
                  </div>
                  <div className="col-md-4">
                    <label className="font">PO Advance Terms</label>
                    <input
                      value={poTerms}
                      onChange={(e) => setPoTerms(e.target.value)}
                      className="form-control"
                    />
                  </div>
                  <div className="col-md-4">
                    <label className="font">PO Amount (GST)</label>
                    <input
                      value={poAmount}
                      onChange={(e) =>
                        handleNumberChange(e.target.value, setPoAmount)
                      }
                      className="form-control"
                    />
                  </div>
                </div>
                <div className="row mb-20">
                  <div className="col-md-4">
                    <label className="font">Request Advance Amount</label>
                    <input
                      value={advanceAmount}
                      onChange={(e) =>
                        handleNumberChange(e.target.value, setAdvanceAmount)
                      }
                      className="form-control"
                    />
                  </div>
                  <div className="col-md-4">
                    <label className="font" style={{ color: "red" }}>
                      Paid Amount
                    </label>
                    <input
                      value={paidAmount}
                      onChange={(e) =>
                        handleNumberChange(e.target.value, setPaidAmount)
                      }
                      className="form-control"
                    />
                  </div>
                  <div className="col-md-4">
                    <label className="font">Expected Settlement Date</label>
                    <input
                      type="date"
                      className="form-control"
                      value={expectedDate}
                      min={new Date().toISOString().split("T")[0]} // 🔥 restrict past dates
                      onChange={(e) => setExpectedDate(e.target.value)}
                    />
                  </div>
                </div>
                <div className="row mb-20">
                  <div className="col-md-4">
                    <label className="font">PIC Name</label>
                    <PeoplePicker
                      context={peoplePickerContext}
                      personSelectionLimit={1}
                      principalTypes={[PrincipalType.User]}
                      defaultSelectedUsers={
                        formData?.PICName?.Title ? [formData.PICName.Title] : []
                      }
                    />
                  </div>
                  <div className="col-md-4">
                    <label className="font">GL Code</label>
                    <input value={glCode} className="form-control readonly" />
                  </div>
                  <div className="col-md-4">
                    <label className="font">Cost Center</label>
                    <input
                      value={employee.CostCenter || ""}
                      className="form-control readonly"
                    />
                  </div>
                </div>
                <div className="row mb-20">
                  <div className="col-md-4">
                    <label className="font">Purpose</label>
                    <textarea
                      value={projectDesc}
                      onChange={(e) => setProjectDesc(e.target.value)}
                      className="form-control"
                    />
                  </div>
                </div>
              </div>
              <div className="heading1" style={{ marginTop: "10px" }}>
                <label>Previous Advances</label>
              </div>
              <div className="main-formcontainer">
                <div className="row mb-20">
                  <div className="col-md-12">
                    <div style={{ overflowX: "auto" }}>
                      <div className="table-vert-scroll">
                        <table className="custom-table min-w-full bg-white rounded-2xl shadow-md">
                          <thead
                            className="text-white"
                            style={{ backgroundColor: "rgb(60, 62, 69)" }}
                          >
                            <tr>
                              <th className="px-4 py-2">PO Number</th>
                              <th className="px-4 py-2">Previous Advance</th>
                              <th className="px-4 py-2">Requested Date</th>
                              <th className="px-4 py-2">Paid Date</th>
                              <th className="px-4 py-2">MRN No</th>
                              <th className="px-4 py-2">Settled Amount</th>
                              <th className="px-4 py-2">Pending Advance</th>
                            </tr>
                          </thead>
                          <tbody>
                            {previousAdvances.length === 0 ? (
                              <tr>
                                <td colSpan={7} style={{ textAlign: "center" }}>
                                  No Data
                                </td>
                              </tr>
                            ) : (
                              previousAdvances.map(
                                (item: any, index: number) => {
                                  const pending = Math.max(
                                    0,
                                    Number(item.RequestAdvanceAmount || 0) -
                                    Number(item.PaidAmount || 0),
                                  );
                                  return (
                                    <tr key={index}>
                                      <td>{item.PONumber}</td>
                                      <td>{item.RequestAdvanceAmount}</td>

                                      <td>
                                        {item.Created
                                          ? new Date(
                                            item.Created,
                                          ).toLocaleDateString()
                                          : ""}
                                      </td>

                                      <td>
                                        {item.VoucherDate
                                          ? new Date(
                                            item.VoucherDate,
                                          ).toLocaleDateString()
                                          : ""}
                                      </td>

                                      <td>{item.VoucherNumber}</td>
                                      <td>{item.PaidAmount}</td>
                                      <td>{pending}</td>
                                    </tr>
                                  );
                                },
                              )
                            )}
                          </tbody>
                        </table>
                      </div>
                    </div>
                  </div>
                </div>
              </div>
              <div className="heading1" style={{ marginTop: "10px" }}>
                <label>Remarks</label>
              </div>
              <div className="main-formcontainer">
                <div className="row mb-20">
                  <div className="col-md-8">
                    <label className="font">Remarks</label>
                    <textarea
                      value={remarks}
                      onChange={(e) => setRemarks(e.target.value)}
                      className="form-control"
                    />
                  </div>
                  <div className="col-md-4">
                    <label className="font">Attachments</label>
                    {attachments.length > 0 && (
                      <ul>
                        {attachments.map((file: any, index: number) => (
                          <li key={index}>
                            <a
                              href={file.ServerRelativeUrl}
                              target="_blank"
                              rel="noopener noreferrer"
                            >
                              {file.Name}
                            </a>

                            {/* 🔥 DELETE FROM SHAREPOINT */}
                            <button
                              type="button"
                              style={{ marginLeft: "10px", color: "red" }}
                              onClick={() => handleDeleteExistingFile(file)}
                            >
                              Delete
                            </button>
                          </li>
                        ))}
                      </ul>
                    )}

                    <input
                      type="file"
                      multiple
                      className="form-control"
                      onChange={(e) => {
                        if (e.target.files) {
                          setSelectedFiles(Array.from(e.target.files));
                        }
                      }}
                    />
                  </div>
                </div>
              </div>
              <div className="heading1" style={{ marginTop: "10px" }}>
                <label>Workflow History</label>
              </div>
              <div className="main-formcontainer">
                <div className="row mb-20">
                  <div className="col-md-12">
                    {workflowHistory.length === 0 ? (
                      <p>No history available</p>
                    ) : (
                      <div className="workflow-history">
                        {/* {workflowHistory.map((h, index) => (
                          <> */}
                            {/* <div key={index} className="history-item">
                            <div>
                              {h.ActionTaken === "Approved" && "✅ "}
                              {h.ActionTaken === "Rejected" && "❌ "}
                              {h.ActionTaken === "Send Back" && "↩ "}
                              {h.ActionTaken}
                            </div>

                            <div>
                              <b>{h.CurrentApprover}</b>
                            </div>
                            <div>{h.Comment}</div>
                            <div className="date">
                              {new Date(h.Date).toLocaleString()}
                            </div>
                          </div> */}

                            <table className="workflow-table" style={{ width: '100%' }}>
                              <thead>
                                <tr>
                                  <th style={{ padding: '8px', textAlign: 'left' }}>Action By</th>
                                  <th style={{ padding: '8px', textAlign: 'left' }}>Action Taken</th>
                                  <th style={{ padding: '8px', textAlign: 'left' }}>Date</th>
                                  <th style={{ padding: '8px', textAlign: 'left' }}>Comment</th>
                                </tr>
                              </thead>
                              <tbody>
                                {workflowHistory.map((h: any, idx: number) => (
                                  <tr key={idx}>
                                    <td style={{ padding: '8px' }}>{h.CurrentApprover || ''}</td>
                                    <td style={{ padding: '8px' }}>{h.ActionTaken || ''}</td>
                                    <td style={{ padding: '8px' }}>{h.Date || ''}</td>
                                    <td style={{ padding: '8px' }}>{h.Comment || ''}</td>
                                  </tr>
                                ))}
                              </tbody>
                            </table>
                          {/* </>
                        ))} */}
                      </div>
                    )}
                  </div>
                </div>
              </div>
              <div
                style={{
                  display: "flex",
                  justifyContent: "center",
                  gap: "10px",
                  margin: "10px",
                }}
              >
                <a onClick={handleSubmit} className="submit-btn">
                  Submit
                </a>
                <a onClick={handledraft} className="Rework-btn">
                  Save as Draft
                </a>
                <a href="#" onClick={handleExit} className="reset-btn">
                  Exit
                </a>
              </div>
            </div>
          </div>
        </div>
      </div>
    </div>
  );
};
export default EditAdvanceForm;
