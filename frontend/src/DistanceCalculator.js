import React, { useState } from "react";
import * as XLSX from "xlsx";
import axios from "axios";
import { Button } from "react-bootstrap";
import Table from "react-bootstrap/Table";
import Form from "react-bootstrap/Form";
import Image from "react-bootstrap/Image";

const DistanceCalculator = () => {
  const [excelData, setExcelData] = useState([]);
  const [results, setResults] = useState([]);

  const downloadTemplate = () => {
    const templateData = [
      {
        "CXP Staff ID": "",
        "OTR Staff ID": "",
        "Staff Name": "",
        "Home Address": "",
        "Store Address": "",
        "Training Location Address": "",
        "Arrive By": "",
        "Depart By": "",
      },
    ];

    const worksheet = XLSX.utils.json_to_sheet(templateData);
    const workbook = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(workbook, worksheet, "Template");

    XLSX.writeFile(workbook, "DistanceCalculatorTemplate.xlsx");
  };
  function formatDate(excelDate) {
    const days = excelDate - 1;

    // Create a new Date object, starting from January 1, 1900
    const jsDate = new Date(1900, 0, days);
    const hours = Math.floor((excelDate - Math.floor(excelDate)) * 24);
    const minutes = Math.round(
      ((excelDate - Math.floor(excelDate)) * 24 - hours) * 60
    );
    jsDate.setHours(hours);
    jsDate.setMinutes(minutes);
    // Format the date and time components
    const year = jsDate.getFullYear();
    const month = (jsDate.getMonth() + 1).toString().padStart(2, "0");
    const day = jsDate.getDate().toString().padStart(2, "0");
    const formattedHours = jsDate.getHours().toString().padStart(2, "0");
    const formattedMinutes = jsDate.getMinutes().toString().padStart(2, "0");

    // Combine the formatted components into the desired format
    const formattedDate = `${day}/${month}/${year} ${formattedHours}:${formattedMinutes}`;

    return formattedDate;
  }
  const exportToExcel = () => {
    const worksheetData = results.map((result) => ({
      "CXP Staff ID": result.cxpStaffId,
      "OTR Staff ID": result.otrStaffId,
      "Staff Name": result.staffName,
      "Home Address": result.home,
      "Store Address": result.store,
      "Training Location Address": result.training,

      "Arrive By": result.arriveBy ? formatDate(result.arriveBy) : "", // Convert to string
      "Depart By": result.departBy ? formatDate(result.departBy) : "", //
      "Distance (Home to Store) (km)": result.homeToStoreDistance,
      "Travel Time (Home to Store) (minutes)": result.homeToStoreTime,
      "Distance (Store to Training) (km)": result.storeToTrainingDistance,
      "Travel Time (Store to Training) (minutes)": result.storeToTrainingTime,
      "Distance (Home to Training) (km)": result.homeToTrainingDistance,
      "Travel Time (Home to Training) (minutes)": result.homeToTrainingTime,
      "Distance (Training to Home) (km)": result.trainingToHomeDistance,
      "Travel Time (Training to Home) (minutes)": result.trainingToHomeTime,
      "Distance (Store to Home) (km)": result.storeToHomeDistance,
      "Travel Time (Store to Home) (minutes)": result.storeToHomeTime,
    }));

    const worksheet = XLSX.utils.json_to_sheet(worksheetData);
    const workbook = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(workbook, worksheet, "Distance Data");
    XLSX.writeFile(workbook, "Distance_Calculations.xlsx");
  };

  const handleFileUpload = (event) => {
    const file = event.target.files[0];
    const reader = new FileReader();

    reader.onload = (e) => {
      const arrayBuffer = e.target.result;
      const workbook = XLSX.read(arrayBuffer, { type: "array" });
      const sheetName = workbook.SheetNames[0];
      const sheet = workbook.Sheets[sheetName];
      const data = XLSX.utils.sheet_to_json(sheet);
      setExcelData(data);
    };

    reader.readAsArrayBuffer(file);
  };

  const calculateDistances = async () => {
    const distances = [];

    for (const row of excelData) {
      const {
        "CXP Staff ID": cxpStaffId,
        "OTR Staff ID": otrStaffId,
        "Staff Name": staffName,
        "Home Address": home,
        "Store Address": store,
        "Training Location Address": training,
        "Arrive By": arriveBy,
        "Depart By": departBy,
      } = row;

      const calculatePair = async (origin, destination, departureTime) => {
        try {
          const response = await axios.get(
            "https://backend-roan-theta.vercel.app/api/",
            {
              params: {
                origins: origin,
                destinations: destination,
                departure_time: departureTime
                  ? new Date(departureTime).getTime() / 1000
                  : undefined,
              },
            }
          );
          const element = response.data.rows[0].elements[0];
          return {
            distance: parseFloat(element.distance.text.replace(/[^\d.-]/g, "")), // Numerical value only
            travelTime: parseFloat(
              element.duration.text.replace(/[^\d.-]/g, "")
            ), // Numerical value only
          };
        } catch (error) {
          console.error(
            `Error fetching distance between ${origin} and ${destination}:`,
            error
          );
          return { distance: 0, travelTime: 0 };
        }
      };

      const homeToStore = await calculatePair(home, store);
      const storeToTraining = await calculatePair(store, training);
      const homeToTraining = await calculatePair(home, training);
      const trainingToHome = await calculatePair(training, home, departBy);
      const storeToHome = await calculatePair(store, home, departBy);

      distances.push({
        cxpStaffId,
        otrStaffId,
        staffName,
        home,
        store,
        training,
        arriveBy, // Include Arrive By
        departBy, // Include Depart By
        homeToStoreDistance: homeToStore.distance,
        homeToStoreTime: homeToStore.travelTime,
        storeToTrainingDistance: storeToTraining.distance,
        storeToTrainingTime: storeToTraining.travelTime,
        homeToTrainingDistance: homeToTraining.distance,
        homeToTrainingTime: homeToTraining.travelTime,
        trainingToHomeDistance: trainingToHome.distance,
        trainingToHomeTime: trainingToHome.travelTime,
        storeToHomeDistance: storeToHome.distance,
        storeToHomeTime: storeToHome.travelTime,
      });
    }

    setResults(distances);
  };

  return (
    <div>
      <h1 style={{ color: "white" }}>OTR Driving Distance Calculator</h1>
      <Button onClick={downloadTemplate} style={{ margin: "5px" }}>
        Download Template
      </Button>
      <div
        style={{
          width: "50%",
          alignSelf: "center",
          justifyItems: "center",
          justifySelf: "center",
          marginTop: "50px",
        }}
      >
        <Form.Group controlId="formFile" className="mb-2 w-5">
          <Form.Control
            type="file"
            accept=".xlsx, .xls"
            onChange={handleFileUpload}
          />
        </Form.Group>
      </div>
      <Button
        onClick={calculateDistances}
        disabled={!excelData.length}
        style={{ margin: "5px" }}
      >
        Calculate Distances
      </Button>
      {results?.length > 0 && (
        <>
          <Button onClick={exportToExcel} style={{ margin: "5px" }}>
            Export to Excel
          </Button>

          <Table striped bordered hover>
            <thead>
              <tr>
                <th>CXP Staff ID</th>
                <th>OTR Staff ID</th>
                <th>Staff Name</th>
                <th>Home Address</th>
                <th>Store Address</th>
                <th>Training Location Address</th>
                <th>Distance (Home to Store) (km)</th>
                <th>Travel Time (Home to Store) (minutes)</th>
                <th>Distance (Store to Training) (km)</th>
                <th>Travel Time (Store to Training) (minutes)</th>
                <th>Distance (Home to Training) (km)</th>
                <th>Travel Time (Home to Training) (minutes)</th>
                <th>Distance (Training to Home) (km)</th>
                <th>Travel Time (Training to Home) (minutes)</th>
                <th>Distance (Store to Home) (km)</th>
                <th>Travel Time (Store to Home) (minutes)</th>
              </tr>
            </thead>
            <tbody>
              {results.map((result, index) => (
                <tr key={index}>
                  <td>{result.cxpStaffId}</td>
                  <td>{result.otrStaffId}</td>
                  <td>{result.staffName}</td>
                  <td>{result.home}</td>
                  <td>{result.store}</td>
                  <td>{result.training}</td>
                  <td>{result.homeToStoreDistance}</td>
                  <td>{result.homeToStoreTime}</td>
                  <td>{result.storeToTrainingDistance}</td>
                  <td>{result.storeToTrainingTime}</td>
                  <td>{result.homeToTrainingDistance}</td>
                  <td>{result.homeToTrainingTime}</td>
                  <td>{result.trainingToHomeDistance}</td>
                  <td>{result.trainingToHomeTime}</td>
                  <td>{result.storeToHomeDistance}</td>
                  <td>{result.storeToHomeTime}</td>
                </tr>
              ))}
            </tbody>
          </Table>
        </>
      )}
    </div>
  );
};

export default DistanceCalculator;
