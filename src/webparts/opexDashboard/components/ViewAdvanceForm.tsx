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
import 'bootstrap/dist/css/bootstrap.min.css';
import logo from "../assets/sona-comstarlogo.png";
interface IVendor {
  Id: number;
  VendorCode: string;
  VendorName: string;
}
const ViewAdvanceForm = ({ context, formData, onClose }: any) => {
  const [attachments, setAttachments] = useState<any[]>([]);
  const sp = spfi().using(SPFx(context));
  const [employee, setEmployee] = useState<any>({});
  // 🔹 Employee
  const [selectedFiles, setSelectedFiles] = useState<File[]>([]);
  const [selectedUser, setSelectedUser] = useState<any[]>([]);
  const [selectedVendorId, setSelectedVendorId] = useState<number | null>(null);
  const [selectedVendorName, setSelectedVendorName] = useState("");

  const [vendors, setVendors] = useState<IVendor[]>([]);
  //const [employeeName, setEmployeeName] = useState("");
  ////const [employeeEmail, setEmployeeEmail] = useState("");
  //const [VendorCode, setVendorCode] = useState<number | null>(null);
  // 🔹 Form fields
  const [poNumber, setPoNumber] = useState("");
  const [poDate, setPoDate] = useState("");
  const [poTerms, setPoTerms] = useState("");
  const [poAmount, setPoAmount] = useState("");
  const [advanceAmount, setAdvanceAmount] = useState("");
  const [paidAmount, setPaidAmount] = useState("");
  const [expectedDate, setExpectedDate] = useState("");
  const [vendorName, setVendorName] = useState("");
  const [glCode, setGlCode] = useState("");
  const [costCenter, setCostCenter] = useState("");
  const [remarks, setRemarks] = useState("");
  const [projectDesc, setProjectDesc] = useState("");

  // 🔹 Extra fields
  const [approverRemarks, setApproverRemarks] = useState("");
  const [voucherDate, setVoucherDate] = useState("");
  const [VouchingNumber, setVouchingNumber] = useState("");
  const [UTRDate, setUTRDate] = useState("");
  const [UTRNumber, setUTRNumber] = useState("");
  const [approvalMatrix, setApprovalMatrix] = useState<any[]>([]);
  const [workflowHistory, setWorkflowHistory] = useState<any[]>([]);

  const peoplePickerContext: IPeoplePickerContext = {
    absoluteUrl: context.pageContext.web.absoluteUrl,
    msGraphClientFactory: context.msGraphClientFactory,
    spHttpClient: context.spHttpClient,
  };
  const getAttachments = async (capexId: string) => {
    try {
      if (!capexId) return;

      const safeCapexId = capexId.replace(/\//g, "_");

      const folderPath = `/sites/SonaFinance/Opex/${safeCapexId}`;

      console.log("Folder Path:", folderPath); // ✅ DEBUG

      const files = await sp.web
        .getFolderByServerRelativePath(folderPath)
        .files();

      console.log("Files:", files); // ✅ DEBUG

      setAttachments(files || []);
    } catch (error) {
      console.log("Attachment fetch error:", error);
      setAttachments([]);
    }
  };

  // const uploadFiles = async () => {
  //     if (!formData?.CapexID || selectedFiles.length === 0) return;

  //     const safe = formData.CapexID.replace(/\//g, "_");
  //     const path = `/sites/SonaFinance/CapexAdvanceDocs/${safe}`;

  //     for (const file of selectedFiles) {
  //       await sp.web
  //         .getFolderByServerRelativePath(path)
  //         .files.addUsingPath(file.name, file, { Overwrite: true });
  //     }

  //   void  setSelectedFiles([]);
  //   void  getAttachments(formData.CapexID);
  //   };

  // ✅ Fetch Item by ID
  const getVendors = async () => {
    const data = await sp.web.lists
      .getByTitle("VendorMaster")
      .items.select("Id", "VendorCode", "VendorName")();
    void setVendors(data);
  };

  const formatDateDDMMYY = (dateStr: string) => {
    const date = new Date(dateStr);
    const dd = String(date.getDate()).padStart(2, "0");
    const mm = String(date.getMonth() + 1).padStart(2, "0");
    const yy = String(date.getFullYear()).slice(-2);
    return `${dd}-${mm}-${yy}`;
  };

  // ✅ Bind SharePoint Data
  useEffect(() => {
    if (!formData) return;

    setPoNumber(formData.PONumber || "");
    setPoDate(formData.PODate?.split("T")[0] || "");
    setPoTerms(formData.POAdvanceTerms || "");
    setPoAmount(formData.POAmtGST || "");
    setAdvanceAmount(formData.RequestAdvanceAmount || "");
    setPaidAmount(formData.PaidAmount || "");
    setExpectedDate(formData.ExpectedDateofSettlement?.split("T")[0] || "");

    setVendorName(formData.VendorName || "");
    setSelectedVendorId(formData.VendorCodeId || null); // ✅ ADD THIS
    setSelectedVendorName(formData.VendorName || ""); // ✅ ADD THIS

    setGlCode(formData.GL || "");
    setCostCenter(formData.CostCenter || "");
    setRemarks(formData.Remarks || "");
    setProjectDesc(formData.Purpose || "");

    setApproverRemarks(formData.ApproverRemarks || "");
    setVoucherDate(formData.VoucherDate?.split("T")[0] || "");
    setVouchingNumber(formData.VouchingNumber || "");
    setUTRDate(formData.UTRDate?.split("T")[0] || "");
    setUTRNumber(formData.UTRNumber || "");

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
    if (formData?.ApprovalMatrix) {
      try {
        const parsed =
          typeof formData.ApprovalMatrix === "string"
            ? JSON.parse(formData.ApprovalMatrix)
            : formData.ApprovalMatrix;

        setApprovalMatrix(Array.isArray(parsed) ? parsed : []);
      } catch (e) {
        console.error("ApprovalMatrix parse error", e);
        setApprovalMatrix([]);
      }
    } else {
      setApprovalMatrix([]);
    }
    // ✅ Workflow History
    if (formData?.WorkFlowHistory) {
      try {
        const parsed =
          typeof formData.WorkFlowHistory === "string"
            ? JSON.parse(formData.WorkFlowHistory)
            : formData.WorkFlowHistory;

        setWorkflowHistory(Array.isArray(parsed) ? parsed : []);
      } catch (e) {
        console.error("WorkFlowHistory parse error", e);
        setWorkflowHistory([]);
      }
    } else {
      setWorkflowHistory([]);
    }
  }, [formData]);

  const handleExit = () => {
    if (onClose) onClose();
    else window.location.reload();
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
  useEffect(() => {
    void getLoggedInUser();
    void getVendors();
  }, []);
  return (
    <div className='MainUplodForm' style={{ margin: "5px 0px" }}>
      <div className='row'>
        <div className='col-md-12'>
          <div className='Main-Boxpoup'>
            {/* 🔹 Header */}
            <div className="bordered">
              <img src={logo} />
              <h1> Advance Payment (View) </h1>
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
                          : a.Status === "Rejected"
                            ? "rejected"
                            : a.Status === "Send Back"
                              ? "sendback"
                              : ""
                        }`}
                    >
                      {a.Role} - {a.Name}
                    </li>
                  ))}
                </ul>
              </div>
            )}

            <div className='borderedbox'>
              <div className="heading1">
                <label>Requestor Information</label>
              </div>
              <div className='main-formcontainer'>
                <div className='row mb-20'>
                  <div className='col-md-4'>
                    <label htmlFor="Employee Code" className='font'>Employee Code</label> : &nbsp;&nbsp;
                    <label className='fonttext'>  {employee.EmployeeCode}</label>
                  </div>
                  <div className='col-md-4'>
                    <label htmlFor="Employee Name" className='font'>Employee Name </label> : &nbsp;&nbsp;
                    <label className='fonttext'>  {employee.EmployeeName}</label>
                  </div>
                  <div className='col-md-4'>
                    <label htmlFor="Employee Email" className='font'>Employee Email </label> : &nbsp;&nbsp;
                    <label className='fonttext'>  {employee.EmployeeEmail}</label>
                  </div>
                </div>
                <div className='row mb-20'>
                  <div className='col-md-4'>
                    <label htmlFor="Contact No" className='font'>Contact No</label> : &nbsp;&nbsp;
                    <label className='fonttext'>  {employee.ContactNo}</label>
                  </div>
                  <div className='col-md-4'>
                    <label htmlFor="Employee Status" className='font'>Employee Status</label> : &nbsp;&nbsp;
                    <label className='fonttext'>  {employee.EmployeeStatus}</label>
                  </div>
                  <div className='col-md-4'>
                    <label htmlFor="Division" className='font'>Division</label> : &nbsp;&nbsp;
                    <label className='fonttext'>  {employee.Division}</label>
                  </div>
                </div>
                <div className='row mb-20'>
                  <div className='col-md-4'>
                    <label htmlFor="Location" className='font'>Location</label> : &nbsp;&nbsp;
                    <label className='fonttext'>  {employee.Location}</label>
                  </div>
                  <div className='col-md-4'>
                    <label htmlFor="RM" className='font'>RM</label> : &nbsp;&nbsp;
                    <label className='fonttext'>  {employee.ReportingManager?.Title}</label>
                  </div>
                  <div className='col-md-4'>
                    <label htmlFor="HOD" className='font'>HOD</label> : &nbsp;&nbsp;
                    <label className='fonttext'>  {employee.HOD?.Title}</label>
                  </div>
                </div>
              </div>
              <div className="heading1" style={{ marginTop: "10px" }}>
                <label>Opex Details</label>
              </div>
              <div className='main-formcontainer'>
                <div className='row mb-20'>
                  <div className='col-md-4'>
                    <label className="font">Vendor Code</label>
                    <select
                      value={selectedVendorId ?? ""}
                      className="form-control readonly"
                      onChange={(e) => {
                        const id = Number(e.target.value);
                        const vendor = vendors.find((v) => v.Id === id);
                        setSelectedVendorId(id);
                        setSelectedVendorName(vendor?.VendorName || "");
                      }}
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
                    <label className="font">Vendor Name</label> : &nbsp;&nbsp;
                    <input type="text" value={vendorName} className="fonttext form-control readonly" />
                  </div>
                  <div className="col-md-4">
                    <label className="font">PO Number</label> : &nbsp;&nbsp;
                    <input type="text" value={poNumber} className="fonttext form-control readonly" />
                  </div>
                </div>
                <div className='row mb-20'>
                  <div className='col-md-4'>
                    <label className="font">PO Date</label> : &nbsp;&nbsp;
                    <input type="date" value={poDate} className="fonttext form-control readonly" />
                  </div>
                  <div className='col-md-4'>
                    <label className="font">PO Terms</label> : &nbsp;&nbsp;
                    <input type="text" value={poTerms} className="fonttext form-control readonly" />
                  </div>
                  <div className="col-md-4">
                    <label className="font">PO Amount</label> : &nbsp;&nbsp;
                    <input type="text" value={poAmount} className="fonttext form-control readonly" />
                  </div>
                </div>
                <div className='row mb-20'>
                  <div className="col-md-4">
                    <label className="font">Advance Amount</label> : &nbsp;&nbsp;
                    <input type="text" value={advanceAmount} className="fonttext form-control readonly" />
                  </div>
                  <div className="col-md-4">
                    <label className="font">Paid Amount</label> : &nbsp;&nbsp;
                    <input type="text" value={paidAmount} className="fonttext form-control readonly" />
                  </div>
                  <div className="col-md-4">
                    <label className="font">Expected Date</label> : &nbsp;&nbsp;
                    <input type="date" value={expectedDate} className="fonttext form-control readonly" />
                  </div>
                </div>
                <div className='row mb-20'>
                  <div className="col-md-4">
                    <label className="font">PIC Name</label> : &nbsp;&nbsp;
                    <PeoplePicker
                      context={peoplePickerContext}
                      personSelectionLimit={1}
                      disabled={true}
                      principalTypes={[PrincipalType.User]}
                      defaultSelectedUsers={
                        formData?.PICName?.Title ? [formData.PICName.Title] : []
                      }
                    />
                  </div>
                  <div className="col-md-4">
                    <label className="font">GL Code</label> : &nbsp;&nbsp;
                    <input type="text" value={glCode} className="fonttext form-control readonly" />
                  </div>
                  <div className="col-md-4">
                    <label className="font">Cost Center</label> : &nbsp;&nbsp;
                    <input type="text" value={costCenter} className="fonttext form-control readonly" />
                  </div>
                </div>
              </div>
              
              <div className="heading1" style={{ marginTop: "10px" }}>
                <label>Vouching & UTR Details</label>
              </div>
              <div className='main-formcontainer'>
                <div className='row mb-20'>
                  <div className='col-md-6'>
                    <label className="font">Voucher Date</label> : &nbsp;&nbsp;
                    <label className='fonttext'>{voucherDate}</label>
                  </div>
                  <div className='col-md-6'>
                    <label className="font">Voucher Number</label> : &nbsp;&nbsp;
                    <label className='fonttext'>{VouchingNumber}</label>
                  </div>
                </div>
                <div className='row mb-20'>
                  <div className='col-md-6'>
                    <label className="font">UTR Date</label> : &nbsp;&nbsp;
                    <label className='fonttext'>{UTRDate}</label>
                  </div>
                  <div className='col-md-6'>
                    <label className="font">UTR Number</label> : &nbsp;&nbsp;
                    <label className='fonttext'>{UTRNumber}</label>
                  </div>
                </div>
              </div>
              <div className="heading1" style={{ marginTop: "10px" }}>
                <label>Uploaded Document</label>
              </div>
              <div className='main-formcontainer'>
                <div className='row mb-20'>
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
                          </li>
                        ))}
                      </ul>
                    )}
                  </div>
                </div>
              </div>
              <div className="heading1" style={{ marginTop: "10px" }}>
                <label>Workflow History</label>
              </div>
              <div className='main-formcontainer'>
                <div className='row mb-20'>
                  <div className='col-md-12'>
                    {workflowHistory.length === 0 ? (
                      <p>No history available</p>
                    ) : (
                      <div className="workflow-history">
                        {/* {workflowHistory.map((h, index) => (
                          <> */}
                            {/* <div key={index} className="history-item">
                              <div>
                                {h.ActionTaken === "Submitted" && "📩 "}
                                {h.ActionTaken === "Approved" && "✅ "}
                                {h.ActionTaken === "Rejected" && "❌ "}
                                {h.ActionTaken === "Send Back" && "↩ "}
                                {h.ActionTaken === "Vouched" && "💰 "}
                                {h.ActionTaken === "Paid" && "💸 "}
                                {h.ActionTaken}
                              </div>

                              <div><b>{h.CurrentApprover}</b></div>
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
              <div className='row my-3'>
                <div className='col-md-12'>
                  <div className='text-center'>
                    <a href="#" onClick={handleExit} className="reset-btn">
                      Exit
                    </a>
                  </div>
                </div>
              </div>
            </div>
          </div>
        </div>
      </div>
    </div>
  );
};

export default ViewAdvanceForm;
