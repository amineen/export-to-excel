import React from 'react';
import Box  from "@mui/material/Box"
import Table from '@mui/material/Table';
import TableBody from '@mui/material/TableBody';
import TableCell from '@mui/material/TableCell';
import TableContainer from '@mui/material/TableContainer';
import TableHead from '@mui/material/TableHead';
import TableRow from '@mui/material/TableRow';
import Paper from '@mui/material/Paper';
import { substationKpis } from './data';
import { Button } from '@mui/material';
import { generateSubstaionKpiExcelFile } from './excelGenerationService';

const App = () => {
  const generateExcel = ()=>{
    generateSubstaionKpiExcelFile();
  }
  return (
    <Box component="h2">
      <Button variant="contained"
      type="button"
        onClick={generateExcel}
      sx={{textTransform:"none", p:1, m:2}}>Export To Ecel</Button>
     <TableContainer component={Paper}>
      <Table sx={{ minWidth: 650 }} aria-label="simple table">
        <TableHead>
          <TableRow>
            <TableCell>Name</TableCell>
            <TableCell align="right">Quarter</TableCell>
            <TableCell align="right">Installed Capacity(MW)</TableCell>
            <TableCell align="right">Peak Load (MW)</TableCell>
            <TableCell align="right">Customers Served</TableCell>
            <TableCell align="right">Total Generation (MWH)</TableCell>
            <TableCell align="right">Total Operation Time (hrs)</TableCell>
            <TableCell align="right">Customers Interrupted </TableCell>
            <TableCell align="right">Duration of Interruption (hrs) </TableCell>
          </TableRow>
        </TableHead>
        <TableBody>
          {substationKpis.map((data) => (
            <TableRow
              key={data.quarter}
              sx={{ '&:last-child td, &:last-child th': { border: 0 } }}
            >
              <TableCell component="th" scope="row">
                {data.name}
              </TableCell>
              <TableCell align="right">{data.quarter}</TableCell>
              <TableCell align="right">{data.installedCapacity}</TableCell>
              <TableCell align="right">{data.peakLoad}</TableCell>
              <TableCell align="right">{data.totalCustomersServed}</TableCell>
              <TableCell align="right">{data.energyDelivered}</TableCell>
              <TableCell align="right">{data.uptime}</TableCell>
              <TableCell align="right">{data.totalCustomersInterrupted}</TableCell>
              <TableCell align="right">{data.durationOfInterruption}</TableCell>
            </TableRow>
          ))}
        </TableBody>
      </Table>
    </TableContainer>
    </Box>
  )
}

export default App
