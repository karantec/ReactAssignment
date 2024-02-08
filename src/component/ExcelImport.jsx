import { useState } from "react";
import { Button, Container, Typography, Box } from "@mui/material";
import * as XLSX from 'xlsx';
import Table from '@mui/material/Table';
import TableBody from '@mui/material/TableBody';
import TableCell from '@mui/material/TableCell';
import TableContainer from '@mui/material/TableContainer';
import TableHead from '@mui/material/TableHead';
import TableRow from '@mui/material/TableRow';
import Paper from '@mui/material/Paper';

function Excellmport() {
  // onchange states
  const [excelFile, setExcelFile] = useState(null);
  const [typeError, setTypeError] = useState(null);
  // submit state
  const [excelData, setExcelData] = useState(null);

  // onchange event
  const handleFile = (e) => {
    // select file type
    let fileTypes = ['application/vnd.ms-excel', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet', 'text/csv'];
    let selectedFile = e.target.files[0];
    if (selectedFile) {
      // check file is present or not 
      if (selectedFile && fileTypes.includes(selectedFile.type)) {
        setTypeError(null);
        //read the file 
        let reader = new FileReader();
        reader.readAsArrayBuffer(selectedFile);
        reader.onload = (e) => {
          setExcelFile(e.target.result);
        }
      }
      else {
        // if file is not excel type
        setTypeError('Please select only excel file types');
        setExcelFile(null);
      }
    }
    else {
      // check file is present or not
      console.log('Please select your file');
    }
  }

  // submit event
  const handleFileSubmit = (e) => {
    e.preventDefault();
    // if file is present
    if (excelFile !== null) {
      const workbook = XLSX.read(excelFile, { type: 'buffer' });
      // copy the first sheet
      const worksheetName = workbook.SheetNames[0];
      // read the data from the first sheet
      const worksheet = workbook.Sheets[worksheetName];
      // convert the data to json
      const data = XLSX.utils.sheet_to_json(worksheet);
      // set the data to the state
      setExcelData(data.slice(0, 10));
    }
  }
  //download excel file
  // download Excel file
  const downloadExcel = () => {
    // Convert the JSON data to Excel file
    const ws = XLSX.utils.json_to_sheet(excelData);
    // Create a new workbook
    const wb = XLSX.utils.book_new();

    //  Append the worksheet to the workbook
    XLSX.utils.book_append_sheet(wb, ws, "Sheet1");

    // Convert the workbook to array buffer
    const wbout = XLSX.write(wb, { bookType: 'xlsx', type: 'array' });
    //  Create a Blob object for the workbook array buffer
    const blob = new Blob([wbout], { type: 'application/octet-stream' });

    // Extract file name from the uploaded file
    const uploadedFileName = document.querySelector('input[type=file]').files[0].name;
    // Extract file name without extension
    const fileName = uploadedFileName.substring(0, uploadedFileName.lastIndexOf('.')) || uploadedFileName;
    // Create a URL for the Blob object
    const url = URL.createObjectURL(blob);
    // Create a new anchor element
    const a = document.createElement('a');
    //  Set the href and download attributes for the anchor element
    a.href = url;
    // Set the download attribute for the anchor element
    a.download = `${fileName}.xlsx`; // Set the downloaded file name
    a.click();
    URL.revokeObjectURL(url);
  };

  return (
    <Container maxWidth="lg" sx={{ textAlign: 'center', mt: 4 }}>
      <Typography variant="h4" gutterBottom sx={{ backgroundColor: '#f0f0f0', padding: '10px' }}>Upload & View Excel Sheets with Download Button</Typography>

      <Box component="form" sx={{ mt: 2 }} onSubmit={handleFileSubmit}>
        <input type="file" accept=".csv, application/vnd.openxmlformats-officedocument.spreadsheetml.sheet, application/vnd.ms-excel" required onChange={handleFile} />
        <Button type="submit" variant="contained" sx={{ ml: 1 }}>UPLOAD</Button>
        {typeError && (
          <div className="alert alert-danger" role="alert">{typeError}</div>
        )}
      </Box>

      <Container maxWidth="sm">
        <Box className="viewer" sx={{ width: 1 }}>
          <Button onClick={downloadExcel} variant="contained" sx={{ mt: 2 }}>Download Excel</Button>
          {excelData ? (

            <TableContainer component={Paper} sx={{ mt: 2 }}>
              <Table>
                <TableHead>
                  <TableRow>
                    {Object.keys(excelData[0]).map((key) => (
                      <TableCell key={key}>{key}</TableCell>
                    ))}
                  </TableRow>
                </TableHead>
                <TableBody>
                  {excelData.map((individualExcelData, index) => (
                    <TableRow key={index}>
                      {Object.keys(individualExcelData).map((key) => (
                        <TableCell key={key}>{individualExcelData[key]}</TableCell>
                      ))}
                    </TableRow>
                  ))}
                </TableBody>
              </Table>
            </TableContainer>

          ) : (
            <Typography variant="body1">No File is uploaded yet!</Typography>
          )}

        </Box>

      </Container>

    </Container>
  );
}

export default Excellmport;
