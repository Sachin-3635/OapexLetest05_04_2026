import * as React from "react";
import "./advanced.scss";
import { spfi } from "@pnp/sp";
import { SPFx } from "@pnp/sp/presets/all";
import { useEffect, useState } from "react";
import { Link, useHistory } from "react-router-dom";
import {
  PeoplePicker,
  PrincipalType,
} from "@pnp/spfx-controls-react/lib/PeoplePicker";
import { IPeoplePickerContext } from "@pnp/spfx-controls-react/lib/PeoplePicker";

import "bootstrap/dist/css/bootstrap.min.css";
import logo from "../assets/sona-comstarlogo.png";

interface IVendor {
  Id: number;
  VendorCode: string;
  VendorName: string;
}

const NewAdvanceform = ({ context }: any) => {
  const history = useHistory();

  const sp = spfi().using(SPFx(context));
  const [employee, setEmployee] = React.useState<any>({});
  //const [selectedUser, setSelectedUser] = useState<any>(null);
  const [attachments, setAttachments] = useState<File[]>([]);
  const [previousAdvances, setPreviousAdvances] = useState<any[]>([]);

  const [employeeName, setEmployeeName] = React.useState("");
  const [pickerKey, setPickerKey] = React.useState<number>(0);
  const [vendors, setVendors] = useState<IVendor[]>([]);

  const [selectedUser, setSelectedUser] = useState<any[]>([]);
  const [selectedVendorId, setSelectedVendorId] = useState<number | null>(null);
  const [selectedVendorName, setSelectedVendorName] = useState("");

  const [poAmount, setPoAmount] = useState("");
  const [advanceAmount, setAdvanceAmount] = useState("");
  const [paidAmount, setPaidAmount] = useState("");
  const [poNumber, setPoNumber] = useState("");
  const [poDate, setPoDate] = useState("");
  const [poTerms, setPoTerms] = useState("");

  const [expectedDate, setExpectedDate] = useState("");

  const [glCode, setGlCode] = useState("390111001");

  const [costCenter, setCostCenter] = useState("");

  const [remarks, setRemarks] = useState("");
  const [Purpose, setPurpose] = useState("");
  const [approvalMatrix, setApprovalMatrix] = useState<any[]>([]);

  const peoplePickerContext: IPeoplePickerContext = {
    absoluteUrl: context.pageContext.web.absoluteUrl,
    msGraphClientFactory: context.msGraphClientFactory,
    spHttpClient: context.spHttpClient,
  };
  const handleNumberChange = (value: string, setter: any) => {
    // Allow only numbers and decimal (max one dot)
    const regex = /^\d*\.?\d*$/;

    if (regex.test(value)) {
      void setter(value);
    }
  };
  const handleRemoveFile = (index: number) => {
  const updatedFiles = [...attachments];
  updatedFiles.splice(index, 1);
  setAttachments(updatedFiles);
};

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

  const handleExit = () => {
    window.location.href =
      "https://isriglobal.sharepoint.com/sites/SonaFinance/SitePages/OpexAdvancedForm.aspx?page=User";
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
          "ReportingManager/Id",
          "HOD/Title",
          "HOD/Id",
          "ContactNo",
          "EmployeeStatus",
          "CostCenter",
        )
        .expand("ReportingManager", "HOD")
        .filter(`EmployeeEmail eq '${email}'`)
        .top(1)();

      if (user.length > 0) {
        void setEmployee(user[0]);
      }
      buildApprovalPreview(user[0]);
    } catch (error) {
      console.log("Error fetching user:", error);
    }
  };
  const getVendors = async () => {
    const data = await sp.web.lists
      .getByTitle("VendorMaster")
      .items.select("Id", "VendorCode", "VendorName")(); // ✅ Id required

    void setVendors(data);
  };

  const getFinancialYear = () => {
    const today = new Date();
    const year = today.getFullYear();
    const month = today.getMonth() + 1;

    if (month >= 4) {
      // April to March
      return `${year.toString().slice(-2)}-${(year + 1).toString().slice(-2)}`;
    } else {
      return `${(year - 1).toString().slice(-2)}-${year.toString().slice(-2)}`;
    }
  };
  const generateCapexId = async () => {
    try {
      const fy = getFinancialYear(); // 25-26

      const items = await sp.web.lists
        .getByTitle("OpexAdvance")
        .items.select("CapexID", "ID")
        .filter(`startswith(CapexID,'OPX/${fy}/')`)
        .orderBy("ID", false)
        .top(1)();

      let nextNumber = 1;

      if (items.length > 0 && items[0].CapexID) {
        const lastId = items[0].CapexID;

        const parts = lastId.split("/");

        if (parts.length === 3) {
          const lastNumber = parseInt(parts[2]);

          if (!isNaN(lastNumber)) {
            nextNumber = lastNumber + 1;
          }
        }
      }

      const formattedNumber = nextNumber.toString().padStart(5, "0");

      return `OPX/${fy}/${formattedNumber}`;
    } catch (error) {
      alert(error);
      console.error("Capex ID Error:", error);
      return `OPX/${getFinancialYear()}/00001`;
    }
  };
  const uploadAttachments = async (capexId: string) => {
    try {
      if (!attachments || attachments.length === 0) return;

      const safeCapexId = capexId.replace(/\//g, "_");

      const libraryName = "Opex";
      const webUrl = context.pageContext.web.serverRelativeUrl;

      const folderPath = `${webUrl}/${libraryName}/${safeCapexId}`;

      // ✅ Ensure folder
      await sp.web.folders.addUsingPath(`${libraryName}/${safeCapexId}`);

      // ✅ Upload files properly
      for (const file of attachments) {
        await sp.web
          .getFolderByServerRelativePath(folderPath)
          .files.addUsingPath(file.name, file, { Overwrite: true });
      }

      console.log("✅ Files uploaded successfully");
    } catch (error) {
      console.error("❌ Upload error:", error);
    }
  };
  const buildApprovalPreview = async (employee: any) => {
    const flow: any[] = [];

    // RM
    if (employee.ReportingManager?.Title) {
      flow.push({
        Name: employee.ReportingManager.Title,
        Role: "RM",
        Status: "Pending",
      });
    }

    // HOD
    if (employee.HOD?.Title) {
      flow.push({
        Name: employee.HOD.Title,
        Role: "HOD",
        Status: "Pending",
      });
    }

    // 🔥 Matrix from list
    const matrixData = await sp.web.lists
      .getByTitle("CapexApprovalMatrix")
      .items.select("Role/RoleName,Approver/Title")
      .expand("Approver,Role")
      .filter("Status eq 'Active'")
      .orderBy("Level", true)();

    const matrixApprovers = matrixData.map((item: any) => ({
      Name: item.Approver?.Title,
      Role: item.Role?.RoleName,
      Status: "Pending",
    }));

    setApprovalMatrix([...flow, ...matrixApprovers]);
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

    if (!Purpose) {
      errors.push("Please enter Purpose");
    }

    if (!attachments || attachments.length === 0) {
      errors.push("Please upload at least one attachment");
    }

    return errors;
  };

  const handleSubmit = async () => {
    try {
      debugger;
      const errors = validateForm();

      if (errors.length > 0) {
        alert(errors.join("\n")); // 👈 shows exactly like your screenshot
        return;
      }

      const capexId = await generateCapexId();

      // ✅ Validate Vendor

      // ✅ Get Email from PeoplePicker
      const userEmail = selectedUser[0]?.secondaryText;

      if (!userEmail) {
        alert("User email not found");
        return;
      }

      // ✅ Ensure User (FIX ERROR)
      const ensuredUser = await sp.web.ensureUser(userEmail);
      const flow = await buildApprovalFlow();

      const currentApprover = flow.length > 0 ? flow[0].Id : null;

      // 🔥 workflow history
      const wfHistory = [
        {
          CurrentApprover: employee.EmployeeName,
          ActionTaken: "Submitted",
          Comment: remarks,
          Date: new Date().toISOString(),
        },
      ];

      await sp.web.lists.getByTitle("OpexAdvance").items.add({
        Title: capexId,
        CapexID: capexId,

        // Employee
        EmployeeCode: employee.EmployeeCode,
        EmployeeName: employee.EmployeeName,
        Division: employee.Division,
        Location: employee.Location,
        Email: employee.EmployeeEmail,
        RM: employee.ReportingManager?.Title,
        HOD: employee.HOD?.Title,
        ContactNo: employee.ContactNo,
        EmployeeStatus: employee.EmployeeStatus,

        // Vendor (LOOKUP)
        VendorCodeId: selectedVendorId, // ✅ FIX
        VendorName: selectedVendorName,

        // PO
        PONumber: poNumber,
        PODate: poDate ? new Date(poDate) : null,
        POAdvanceTerms: poTerms,

        // Amount
        POAmtGST: poAmount,
        RequestAdvanceAmount: advanceAmount,
        PaidAmount: paidAmount,

        // Advance
        ExpectedDateofSettlement: expectedDate ? new Date(expectedDate) : null,

        // Person field
        PICNameId: ensuredUser.Id, // ✅ FIX

        // Other
        GL: glCode,
        CostCenter: employee.CostCenter,
        Remarks: remarks,
        Purpose: Purpose,

        Status: "Pending for Approver",
        ApprovalMatrix: JSON.stringify(flow),

        CurrentApproverId: currentApprover,

        WorkFlowHistory: JSON.stringify(wfHistory),
        ApproverStatus:"Pending at RM"
      });
      debugger;
      await uploadAttachments(capexId); // 🔥 FIXED

      console.log("Attachments:", attachments);
      alert("Submitted successfully ✅");

      // 🔥 REDIRECT
      window.location.href =
        "https://isriglobal.sharepoint.com/sites/SonaFinance/SitePages/OpexAdvancedForm.aspx?page=User";
    } catch (error) {
      console.error("ERROR:", error);
      alert("Error while saving ❌");
    }
  };

  const handledraft = async () => {
    try {
      const capexId = await generateCapexId();

      let ensuredUserId: number | null = null;

      // ✅ Only process if user selected
      if (selectedUser && selectedUser.length > 0) {
        const userEmail = selectedUser[0]?.secondaryText;

        if (userEmail) {
          const ensuredUser = await sp.web.ensureUser(userEmail);
          ensuredUserId = ensuredUser.Id;
        }
      }
      const flow = await buildApprovalFlow();

      // 🔥 ALL SHOULD BE PENDING (NO IN PROGRESS)
      flow.forEach((f: any) => (f.Status = "Pending"));
      const currentApprover = flow.length > 0 ? flow[0].Id : null;

      // =========================
      // 🔥 WORKFLOW HISTORY (DRAFT)
      // =========================
      const wfHistory = [
        {
          CurrentApprover: employee.EmployeeName,
          ActionTaken: "Draft Saved",
          Comment: remarks || "",
          Date: new Date().toISOString(),
        },
      ];
      await sp.web.lists.getByTitle("OpexAdvance").items.add({
        Title: capexId,
        CapexID: capexId,

        // Employee
        EmployeeCode: employee.EmployeeCode,
        EmployeeName: employee.EmployeeName,
        Division: employee.Division,
        Location: employee.Location,
        Email: employee.EmployeeEmail,
        RM: employee.ReportingManager?.Title,
        HOD: employee.HOD?.Title,
        ContactNo: employee.ContactNo,
        EmployeeStatus: employee.EmployeeStatus,

        // Vendor
        VendorCodeId: selectedVendorId,
        VendorName: selectedVendorName,

        // PO
        PONumber: poNumber,
        PODate: poDate ? new Date(poDate) : null,
        POAdvanceTerms: poTerms,

        // Amount
        POAmtGST: poAmount,
        RequestAdvanceAmount: advanceAmount,
        PaidAmount: paidAmount,

        // Advance
        ExpectedDateofSettlement: expectedDate ? new Date(expectedDate) : null,

        // ✅ PIC (OPTIONAL)
        ...(ensuredUserId && { PICNameId: ensuredUserId }),

        // Other
        GL: glCode,
        CostCenter: employee.CostCenter,
        Remarks: remarks,
        Purpose: Purpose,

        Status: "Draft",
        ApprovalMatrix: JSON.stringify(flow),

        CurrentApproverId: currentApprover, // 🔥 not started

        WorkFlowHistory: JSON.stringify(wfHistory),
      });

      const safeCapexId = capexId.replace(/\//g, "_");
      void uploadAttachments(safeCapexId);

      alert("Draft saved successfully ✅");

      window.location.href =
        "https://isriglobal.sharepoint.com/sites/SonaFinance/SitePages/OpexAdvancedForm.aspx?page=User";
    } catch (error) {
      console.error("ERROR:", error);
      alert("Error while saving ❌");
    }
  };

  React.useEffect(() => {
    if (!context) return;

    void getLoggedInUser();
    void getVendors(); // 👈 ADD THIS
    //  buildApprovalPreview();
  }, [context]);

  return (
    <div className="MainUplodForm" style={{ margin: "5px 0px" }}>
      <div className="row">
        <div className="col-md-12">
          <div className="Main-Boxpoup">
            {/* 🔹 Header */}
            <div className="bordered">
              <img src={logo} />
              <h1> Opex Advance Request </h1>
            </div>
            {approvalMatrix.length === 0 ? (
              <p>Loading...</p>
            ) : (
              <div className="displayWF">
                <ul className="approval-flow">
                  {approvalMatrix.map((a, index) => (
                    <li
                      key={index}
                      className={`approval-step ${index === 0 ? "active" : ""}`}
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
                <label>Vendor & PO Details</label>
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
                  </div>
                  <div className="col-md-4">
                    <label className="font">Vendor Name</label>
                    <input
                      value={selectedVendorName}
                      className="form-control readonly"
                    />
                  </div>
                  <div className="col-md-4">
                    <label className="font">PO Number</label>
                    <input
                      value={poNumber}
                      className="form-control"
                      onChange={(e) => setPoNumber(e.target.value)}
                    />
                  </div>
                </div>
                <div className="row mb-20">
                  <div className="col-md-4">
                    <label className="font">PO Date</label>
                    <input
                      type="date"
                      value={poDate}
                      className="form-control"
                      onChange={(e) => setPoDate(e.target.value)}
                    />
                  </div>
                  <div className="col-md-4">
                    <label className="font">PO Advance Terms</label>
                    <input
                      value={poTerms}
                      className="form-control"
                      onChange={(e) => setPoTerms(e.target.value)}
                    />
                  </div>
                  <div className="col-md-4">
                    <label className="font">PO Amount (GST)</label>
                    <input
                      value={poAmount}
                      className="form-control"
                      onChange={(e) =>
                        handleNumberChange(e.target.value, setPoAmount)
                      }
                    />
                  </div>
                </div>
                <div className="row mb-20">
                  <div className="col-md-4">
                    <label className="font">Request Advance Amount</label>
                    <input
                      value={advanceAmount}
                      className="form-control"
                      onChange={(e) =>
                        handleNumberChange(e.target.value, setAdvanceAmount)
                      }
                    />
                  </div>
                  <div className="col-md-4">
                    <label className="font" style={{ color: "red" }}>
                      Paid Amount
                    </label>
                    <input
                      value={paidAmount}
                      className="form-control"
                      onChange={(e) =>
                        handleNumberChange(e.target.value, setPaidAmount)
                      }
                    />
                  </div>
                </div>
              </div>
              <div className="heading1" style={{ marginTop: "10px" }}>
                <label>Advance Details</label>
              </div>
              <div className="main-formcontainer">
                <div className="row mb-20">
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
                  <div className="col-md-4">
                    <label className="font">PIC Name</label>
                    <PeoplePicker
                      context={peoplePickerContext}
                      personSelectionLimit={1}
                      ensureUser={true}
                      principalTypes={[PrincipalType.User]}
                      onChange={(items) => setSelectedUser(items)}
                    />
                  </div>
                  <div className="col-md-4">
                    <label className="font">GL Code</label>
                    <input value={glCode} className="form-control readonly" />
                  </div>
                </div>
                <div className="row mb-20">
                  <div className="col-md-4">
                    <label className="font">Cost Center</label>
                    <input
                      value={employee.CostCenter}
                      className="form-control readonly"
                    />
                  </div>
                </div>
              </div>
              <div className="heading1" style={{ marginTop: "10px" }}>
                <label>Remarks</label>
              </div>
              <div className="main-formcontainer">
                <div className="row mb-20">
                  <div className="col-md-4">
                    <label className="font">Remarks</label>
                    <textarea
                      value={remarks}
                      className="font-control"
                      onChange={(e) => setRemarks(e.target.value)}
                    />
                  </div>
                  <div className="col-md-4">
                    <label className="font">Purpose</label>
                    <textarea
                      value={Purpose}
                      className="font-control"
                      onChange={(e) => setPurpose(e.target.value)}
                    />
                  </div>
                  <div className="col-md-4">
                    <div className="col-md-4">
                      <label className="font">Attach</label>

                      <input
                        type="file"
                        multiple
                        onChange={(e) => {
                          if (e.target.files) {
                            setAttachments((prev) => [
                              ...prev,
                              ...Array.from(e.target.files),
                            ]);
                          }
                        }}
                      />

                      {/* 🔥 Show selected files */}
                      {attachments.length > 0 && (
                        <ul style={{ marginTop: "10px" }}>
                          {attachments.map((file, index) => (
                            <li key={index}>
                              {file.name}

                              {/* 🔥 Remove button */}
                              <button
                                type="button"
                                style={{
                                  marginLeft: "10px",
                                  color: "red",
                                  cursor: "pointer",
                                }}
                                onClick={() => handleRemoveFile(index)}
                              >
                                Remove
                              </button>
                            </li>
                          ))}
                        </ul>
                      )}
                    </div>
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
              <div
                style={{
                  display: "flex",
                  justifyContent: "center",
                  gap: "5px",
                  marginBottom: "1rem",
                  marginTop: "1rem",
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

export default NewAdvanceform;
