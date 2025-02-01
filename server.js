const path = require('path');
const express = require('express');
const fs = require('fs');
const xlsx = require('xlsx');
const { format, parse } = require('date-fns');
const bodyParser = require('body-parser');

function convertDate(inputDate) {
  if (!inputDate) return '';

  const inputFormat = 'yyyy-MM-dd';
  const outputFormat = 'dd/MM/yyyy';
  const parsedDate = parse(inputDate, inputFormat, new Date());

  // console.log(parsedDate, 'parsed date');

  const formattedDate = format(parsedDate, outputFormat);
  return formattedDate;
}

const app = express();
const PORT = 4001;

// app.use(express.json());
app.use(bodyParser.json({ limit: '50mb' }));
app.use(bodyParser.urlencoded({ limit: '50mb', extended: true }));

const fileDir = path.join(__dirname, 'files');
if (!fs.existsSync(fileDir)) {
  fs.mkdirSync(fileDir);
}

app.post('/create-excel', (req, res) => {
  const data = req.body;

  const refactoredData = data?.map((item, i) => {
    const data = {
      Roll_no: item?.rollNo,
      Email: item?.email ? item?.email?.trim()?.toLowerCase() : '',
      password: item?.password,
      Student_name: item?.studentName,
      Gender: item?.gender,
      Date_of_birth: convertDate(item?.DOB),
      Aadhar_no: item?.AadharNo,
      Blood_Group: item?.bloodGp,
      Contact_no_2: item?.AltCnctNo,
      Standard: item?.std,
      Admission_no: item?.admnNo,
      Contact_no: item?.ContactNo,
      Academic_year: item?.academicYear,
      Academic_Year_Of_Join: item?.academicYearOfJoin,
      Height_in_Cm: item?.height,
      Weight_in_Kg: item?.weight,
      Mother_tongue: item?.motherTongue,
      Pincode: item?.Pincode, // Ignored defaultPinCode
      City: item?.city,
      District: item?.district,
      State: item?.state,
      Nationality: item?.nationality,
      Religion: item?.religion,
      Community: item?.community,
      Caste: item?.caste,
      SubCaste: item?.subCaste,
      Address: item?.address,

      Father_name: item?.FathersName,
      Father_occupation: item?.FathersJob,
      Father_contact_no: item?.FathersPhNo,
      Father_work_address: item?.FathersWorkAddress,
      Father_email: item?.FathersMailId,
      Father_annual_income: item?.fatherannualIncome,

      Mother_name: item?.MothersName,
      Mother_occupation: item?.MothersJob,
      Mother_contact_no: item?.MothersPhNo,
      Mother_work_address: item?.MothersWorkAddress,
      Mother_email: item?.MothersMailId,
      Mother_annual_income: item?.motherannualIncome,

      Guardian_name: item?.guardianName,
      Guardian_occupation: item?.guardiansJob,
      Guardian_contact_no: item?.guardianPhNo,
      Guardian_work_address: item?.guardianWorkAddress,
      Guardian_email: item?.guardianMailId,
      Guardian_annual_income: item?.guardianannualIncome,

      Class_of_joining: item?.classOfJoin,
      Year_of_joining: item?.yearOfJoin,
      EMIS_no: item?.EMISno,
      Selected_group: item?.studentGp,
      Medium_of_instruction: item?.mediumOfInstruction,
      Emergency_contact_no: item?.emergency_contactNo,
      Mode_of_study: item?.modeOfStudy,
      Eligibility_Certificate_No: item?.eligibilityCertificateNo,
      Migration_Certificate_No: item?.migration_certificate_no,
      // Migration_Certificate_Date: convertDate(item?.migration_certificate_date),
      // Migration_Certificate_Issue_Authority: convertDate(item?.migration_certificate_issue_authority),
      Community_Certificate_No: item?.community_certificate_no,
      // Community_Certificate_Date: convertDate(item?.community_certificate_date),
      Community_Certificate_Issue_Authority: item?.community_certificate_issue_authority,
      Village_Taluk: item?.village_taluk,

      _10thPassingMonth: item?._10thPassingMonth,
      _10thRegNo: item?._10thRegNo,
      _10th_Certificate_No: item?._10thCertificateNo,
      // _10th_Certificate_Date: convertDate(item?._10thCertificateDate),
      _10th_Issue_authority: item?._10thIssue_authority,
      _12th_Passing_Month_and_Year: item?._12thPassingMonth,
      _12th_Reg_No: item?._12thRegNo,
      _12th_Certificate_No: item?._12thCertificateNo,
      // _12th_Certificate_Date: convertDate(item?._12thCertificateDate),
      _12th_Issue_authority: item?._12thIssue_authority,

      Quota: item?.selectedQuota, // Check if this is correct
    };
    return data;
  });

  const worksheet = xlsx.utils.json_to_sheet(refactoredData);

  // Create a new workbook and append the worksheet
  const workbook = xlsx.utils.book_new();
  xlsx.utils.book_append_sheet(workbook, worksheet, 'Sheet1');

  // Generate a unique filename
  const filename = `data_${Date.now()}.xlsx`;
  const filePath = path.join(fileDir, filename);

  // Write the workbook to the file
  xlsx.writeFile(workbook, filePath);

  res.send(`Excel file created successfully: ${filename}`);
});

app.listen(PORT, () => console.log(`Server running on PORT ${PORT}`));

const mockSchema = {
  rollNo: String,
  email: String,
  password: String,
  studentName: String,
  gender: String,
  DOB: String,
  AadharNo: String,
  bloodGp: String,
  AltCnctNo: String,
  std: String,
  admnNo: String,
  ContactNo: String,
  academicYear: String,
  academicYearOfJoin: String,
  height: String,
  weight: String,
  motherTongue: String,
  pincode: String,
  city: String,
  district: String,
  state: String,
  nationality: String,
  religion: String,
  community: String,
  caste: String,
  subCaste: String,
  address: String,

  FathersName: String,
  FathersJob: String,
  FathersPhNo: String,
  FathersWorkAddress: String,
  FathersMailId: String,
  fatherannualIncome: String,

  MothersName: String,
  MothersJob: String,
  MothersPhNo: String,
  MothersWorkAddress: String,
  MothersMailId: String,
  motherannualIncome: String,

  guardianName: String,
  guardiansJob: String,
  guardianPhNo: String,
  guardianWorkAddress: String,
  guardianMailId: String,
  guardianannualIncome: String,

  classOfJoin: String,
  yearOfJoin: String,
  EMISno: String,
  studentGp: String,
  mediumOfInstruction: String,
  emergency_contactNo: String,
  modeOfStudy: String,
  eligibilityCertificateNo: String,
  migration_certificate_no: String,
  migration_certificate_date: String,
  migration_certificate_issue_authority: String,
  community_certificate_no: String,
  community_certificate_date: String,
  community_certificate_issue_authority: String,
  village_taluk: String,

  _10thPassingMonth: String,
  _10thRegNo: String,
  _10thCertificateNo: String,
  _10thCertificateDate: String,
  _10thIssue_authority: String,

  _12thPassingMonth: String,
  _12thRegNo: String,
  _12thCertificateNo: String,
  _12thCertificateDate: String,
  _12thIssue_authority: String,

  selectedQuota: String,
};
