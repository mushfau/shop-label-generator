/* eslint-disable @typescript-eslint/no-explicit-any */
"use client";

import { useState } from 'react';
import * as XLSX from 'xlsx';
import { jsPDF } from 'jspdf';
import { createErasLight } from './ERASLGHT-normal'
import { createErasBold } from './ERASBD-bold'
jsPDF.API.events.push(['addFonts', createErasLight])
jsPDF.API.events.push(['addFonts', createErasBold])


function formatFinancial(number: any) {
  // Handle invalid inputs
  if (number === null || number === undefined || isNaN(number)) {
    return '0.00';
  }

  // Convert to number if it's a string
  const num = typeof number === 'string' ? parseFloat(number) : number;

  // Format with 2 decimal places and thousand separators
  return num.toLocaleString('en-US', {
    minimumFractionDigits: 2,
    maximumFractionDigits: 2
  });
}

function formatCurrency(number: number, symbol = '') {
  const formatted = formatFinancial(number);
  return `${symbol} ${formatted}`.trim();
}

export default function ExcelToPDFLabels() {
  const [file, setFile] = useState<any>(null);
  const [data, setData] = useState<any>([]);
  const [preview, setPreview] = useState<any>([]);
  const [isLoading, setIsLoading] = useState(false);
  const [error, setError] = useState("");
  const [labelSettings, setLabelSettings] = useState({
    currency: '',
    labelsPerPage: 12,
    pageWidth: 210,
    pageHeight: 297,
    labelWidth: 50,
    labelHeight: 25,
    marginTop: 10,
    marginLeft: 10,
    columnGap: 10,
    rowGap: 10,
    columns: 3
  });

  const downloadSampleFile = () => {

    const sampleData = [
      { Price: 1095.50, Description: "Stainless Steel Bollard Center Curved 10mm x 70mm Left", Code: "10148", LabelCount: 2 },
      { Price: 95.50, Description: "Stainless Steel Bollard Center Curved 10mm x 70mm Right", Code: "10149", LabelCount: 1 },
      { Price: 2278.10, Description: "Stainless Steel Bollard Center Curved 12mm x 120mm Right", Code: "10957", LabelCount: 3 },
      { Price: 125.00, Description: "Stainless Steel Handle 200mm", Code: "10224", LabelCount: 1 },
      { Price: 45.75, Description: "Aluminum Bracket Small", Code: "10332", LabelCount: 4 },
    ];


    try {
      // Create a new workbook
      const workbook = XLSX.utils.book_new();

      // Convert sample data to worksheet
      const worksheet = XLSX.utils.json_to_sheet(sampleData);

      // Add the worksheet to the workbook
      XLSX.utils.book_append_sheet(workbook, worksheet, "Labels");

      // Generate Excel file and trigger download
      XLSX.writeFile(workbook, "sample-labels.xlsx");
    } catch (err) {
      setError("Error generating sample file.");
      console.error("Sample file generation error:", err);
    }
  };

  const handleFileUpload = (e: any) => {
    const file = e.target.files[0];
    setFile(file);
    setError("");

    if (file) {
      const reader = new FileReader();
      reader.onload = (evt: any) => {
        try {
          const binaryData = evt.target.result;
          const workbook = XLSX.read(binaryData, { type: 'binary' });
          const firstSheetName = workbook.SheetNames[0];
          const worksheet = workbook.Sheets[firstSheetName];
          const jsonData = XLSX.utils.sheet_to_json(worksheet);

          setData(jsonData);
          setPreview(jsonData.slice(0, labelSettings.labelsPerPage)); // Preview first 6 items

          // Calculate total labels
          let totalLabels = 0;
          jsonData.forEach((item: any) => {
            const labelCount = parseInt(item.LabelCount || item.labelcount || item.Quantity || item.quantity || '1', 10);
            totalLabels += labelCount;
          });

          // Show info about total labels
          if (totalLabels > jsonData.length) {
            setError(`Info: ${jsonData.length} records will generate ${totalLabels} labels based on LabelCount values.`);
          }
        } catch (err) {
          setError("Error processing Excel file. Please ensure it's a valid Excel format.");
          console.error("Error reading Excel:", err);
        }
      };
      reader.onerror = () => {
        setError("Error reading file. Please try again.");
      };
      reader.readAsBinaryString(file);
    }
  };

  const generatePDF = () => {
    if (!data.length) {
      setError("No data available. Please upload an Excel file first.");
      return;
    }

    setIsLoading(true);

    try {
      // PDF setup - A4 size in mm (210 x 297)
      const doc = new jsPDF({
        orientation: 'portrait',
        unit: 'mm',
        format: 'a4'
      });

      const {
        labelWidth, labelHeight, marginTop, marginLeft,
        columnGap, rowGap, columns, labelsPerPage, currency
      } = labelSettings;

      // const rows = Math.floor(labelsPerPage / columns);

      // Expand data based on LabelCount
      const expandedData: any = [];
      data.forEach((item: any) => {
        const labelCount = parseInt(item.LabelCount || item.labelcount || item.Quantity || item.quantity || '1', 10);
        for (let i = 0; i < labelCount; i++) {
          expandedData.push(item);
        }
      });

      // Process each item and create labels
      expandedData.forEach((item: any, index: number) => {
        // // Calculate page number and position
        // // const page = Math.floor(index / labelsPerPage);
        // const positionOnPage = index % labelsPerPage;
        // const row = Math.floor(positionOnPage / columns);
        // const col = positionOnPage % columns;

        // // Add new page if needed
        // if (positionOnPage === 0 && index > 0) {
        //   doc.addPage();
        // }

        // // Calculate x and y position for this label
        // const x = marginLeft + (col * (labelWidth + columnGap));
        // const y = marginTop + (row * (labelHeight + rowGap));

        // // Draw label background
        // doc.setFillColor(240, 240, 240);
        // doc.rect(x, y, labelWidth, labelHeight, 'F');

        // // Draw border
        // doc.setDrawColor(200, 200, 200);
        // doc.rect(x, y, labelWidth, labelHeight, 'S');

        // // Draw price at top in large font - centered
        // doc.setFontSize(12);
        // doc.setFont('helvetica', 'normal');
        // const price = item.Price.toLocaleString(undefined, { maximumFractionDigits: 2, minimumFractionDigits: 2 }) || item.price.toLocaleString(undefined, { maximumFractionDigits: 2, minimumFractionDigits: 2 }) || item.PRICE.toLocaleString(undefined, { maximumFractionDigits: 2, minimumFractionDigits: 2 }) || '0.00';
        // doc.text(`${price}`, x + (labelWidth / 2), y + 10, { align: 'center' });

        // // Draw description text - centered
        // doc.setFontSize(10);
        // doc.setFont('helvetica', 'normal');

        // const description = item.Description || item.description || item.DESC || item.Name || item.name || '';
        // const descriptionLines = doc.splitTextToSize(description, labelWidth - 10);
        // doc.text(descriptionLines, x + (labelWidth / 2), y + 20, { align: 'center' });

        // // Draw item number/code at bottom - centered
        // doc.setFontSize(12);
        // doc.setFont('helvetica', 'normal');
        // const itemCode = item.Code || item.code || item.ID || item.id || item.ItemNumber || `${index + 1}`;
        // doc.text(`${itemCode}`, x + (labelWidth / 2), y + labelHeight - 5, { align: 'center' });


        const positionOnPage = index % labelsPerPage;
        const row = Math.floor(positionOnPage / columns);
        const col = positionOnPage % columns;

        // Add new page if needed
        if (positionOnPage === 0 && index > 0) {
          doc.addPage();
        }

        // Calculate x and y position for this label
        const x = marginLeft + (col * (labelWidth + columnGap));
        const y = marginTop + (row * (labelHeight + rowGap));

        // Draw label background
        doc.setFillColor(255, 255, 255);
        doc.rect(x, y, labelWidth, labelHeight, 'F');

        // Draw border
        doc.setDrawColor(200, 200, 200);
        doc.rect(x, y, labelWidth, labelHeight, 'S');

        // Draw price at top in large font - centered
        doc.setFontSize(16);
        doc.setFont('ERASBD', 'bold');
        const price = formatCurrency(item.Price, currency) || formatCurrency(item.price, currency) || formatCurrency(item.PRICE, currency) || '0.00';
        doc.text(`${price}`, x + (labelWidth / 2), y + 6, { align: 'center' });

        // Draw description text - vertically and horizontally centered
        doc.setFontSize(10);
        doc.setFont('ERASLGHT', 'normal');

        const description = item.Description || item.description || item.DESC || item.Name || item.name || '';
        const descriptionLines = doc.splitTextToSize(description, labelWidth - 10);

        // Calculate vertical center position
        const textHeight = descriptionLines.length * 3.5; // Approximate line height
        const textStartY = y + (labelHeight / 2) - (textHeight / 2) + 3;

        doc.text(descriptionLines, x + (labelWidth / 2), textStartY, { align: 'center' });

        // Draw item number/code at bottom - centered
        doc.setFontSize(12);
        doc.setFont('ERASBD', 'bold');
        const itemCode = item.Code || item.code || item.ID || item.id || item.ItemNumber || `${index + 1}`;
        doc.text(`${itemCode}`, x + (labelWidth / 2), y + labelHeight - 2, { align: 'center' });


      });

      // Save the PDF
      doc.save("product-labels.pdf");
      setIsLoading(false);
    } catch (err) {
      setError("Error generating PDF. Please try again.");
      setIsLoading(false);
      console.error("PDF generation error:", err);
    }
  };

  const handleSettingsChange = (e: any) => {
    const { name, value } = e.target;
    setLabelSettings({
      ...labelSettings,
      [name]: parseFloat(value)
    });
  };

  const handleCurrencyChange = (e: any) => {
    const { name, value } = e.target;
    setLabelSettings({
      ...labelSettings,
      [name]: value
    });
  };

  return (
    <div className="flex flex-col p-6 max-w-4xl mx-auto bg-white rounded-lg shadow-md">
      <h1 className="text-2xl font-bold mb-4 text-gray-800">Excel to PDF Label Generator</h1>

      <div className="mb-6 p-4 bg-blue-50 rounded-lg">
        <h2 className="text-lg font-semibold mb-2">Instructions</h2>
        <ol className="list-decimal pl-5 space-y-1 text-sm">
          <li>Upload an Excel file with your product data</li>
          <li>Adjust label settings if needed</li>
          <li>Click `Generate PDF` to create your labels</li>
          <li>Download the generated PDF file</li>
        </ol>
        <p className="mt-2 text-sm text-gray-600">
          Your Excel file should include columns for:
        </p>
        <ul className="list-disc pl-5 space-y-1 text-sm">
          <li><strong>Price</strong>: The price displayed at the top of the label</li>
          <li><strong>Description</strong>: The product description shown in the middle</li>
          <li><strong>Code</strong>: The item code displayed at the bottom</li>
          <li><strong>LabelCount</strong>: How many copies of each label to print (defaults to 1 if not specified)</li>
        </ul>
        <div className="mt-3">
          <button
            onClick={downloadSampleFile}
            className="px-3 py-1 text-sm bg-blue-600 text-white rounded hover:bg-blue-700 focus:outline-none focus:ring-2 focus:ring-blue-500"
          >
            Download Sample Excel File
          </button>
        </div>
      </div>

      <div className="mb-6">
        <label className="block text-sm font-medium text-gray-700 mb-2">
          Upload Excel File:
        </label>
        <input
          type="file"
          accept=".xlsx, .xls"
          onChange={handleFileUpload}
          className="block w-full text-sm text-gray-500 file:mr-4 file:py-2 file:px-4 file:rounded-md file:border-0 file:text-sm file:font-semibold file:bg-blue-50 file:text-blue-700 hover:file:bg-blue-100"
        />
        {file && <p className="mt-1 text-sm text-gray-500">File: {file.name}</p>}
      </div>

      {error && (
        <div className={`mb-4 p-3 ${error.startsWith('Info:') ? 'bg-blue-50 text-blue-700' : 'bg-red-50 text-red-700'} rounded-md`}>
          {error}
        </div>
      )}

      {preview.length > 0 && (
        <div className="mb-6">
          <h2 className="text-lg font-semibold mb-2">Data Preview</h2>
          <div className="overflow-auto h-52 ">
            <table className="min-w-full divide-y divide-gray-200">
              <thead className="bg-gray-50">
                <tr>
                  {Object.keys(preview[0]).map(key => (
                    <th key={key} className="px-6 py-3 text-left text-xs font-medium text-gray-500 uppercase tracking-wider">
                      {key}
                    </th>
                  ))}
                </tr>
              </thead>
              <tbody className="bg-white divide-y divide-gray-200">
                {preview.map((row: any, idx: number) => (
                  <tr key={idx}>
                    {Object.values(row).map((val: any, i) => (
                      <td key={i} className="px-6 py-4 whitespace-nowrap text-sm text-gray-500">
                        {val.toString()}
                      </td>
                    ))}
                  </tr>
                ))}
              </tbody>
            </table>
          </div>
          <p className="mt-1 text-sm text-gray-500">
            Showing {preview.length} of {data.length} items
          </p>
        </div>
      )}

      <div className="mb-6">
        <h2 className="text-lg font-semibold mb-2">Label Settings</h2>
        <div className="grid grid-cols-1 md:grid-cols-2 gap-4">
          <div>
            <label className="block text-sm font-medium text-gray-700 mb-1">
              Labels Per Page:
            </label>
            <input
              type="number"
              name="labelsPerPage"
              value={labelSettings.labelsPerPage}
              onChange={handleSettingsChange}
              className="mt-1 block w-full rounded-md border-gray-300 shadow-sm focus:border-blue-500 focus:ring-blue-500 sm:text-sm"
            />
          </div>
          <div>
            <label className="block text-sm font-medium text-gray-700 mb-1">
              Columns:
            </label>
            <input
              type="number"
              name="columns"
              value={labelSettings.columns}
              onChange={handleSettingsChange}
              className="mt-1 block w-full rounded-md border-gray-300 shadow-sm focus:border-blue-500 focus:ring-blue-500 sm:text-sm"
            />
          </div>
          <div>
            <label className="block text-sm font-medium text-gray-700 mb-1">
              Label Width (mm):
            </label>
            <input
              type="number"
              name="labelWidth"
              value={labelSettings.labelWidth}
              onChange={handleSettingsChange}
              className="mt-1 block w-full rounded-md border-gray-300 shadow-sm focus:border-blue-500 focus:ring-blue-500 sm:text-sm"
            />
          </div>
          <div>
            <label className="block text-sm font-medium text-gray-700 mb-1">
              Label Height (mm):
            </label>
            <input
              type="number"
              name="labelHeight"
              value={labelSettings.labelHeight}
              onChange={handleSettingsChange}
              className="mt-1 block w-full rounded-md border-gray-300 shadow-sm focus:border-blue-500 focus:ring-blue-500 sm:text-sm"
            />
          </div>
          <div>
            <label className="block text-sm font-medium text-gray-700 mb-1">
              Margin Top (mm):
            </label>
            <input
              type="number"
              name="marginTop"
              value={labelSettings.marginTop}
              onChange={handleSettingsChange}
              className="mt-1 block w-full rounded-md border-gray-300 shadow-sm focus:border-blue-500 focus:ring-blue-500 sm:text-sm"
            />
          </div>
          <div>
            <label className="block text-sm font-medium text-gray-700 mb-1">
              Margin Left (mm):
            </label>
            <input
              type="number"
              name="marginLeft"
              value={labelSettings.marginLeft}
              onChange={handleSettingsChange}
              className="mt-1 block w-full rounded-md border-gray-300 shadow-sm focus:border-blue-500 focus:ring-blue-500 sm:text-sm"
            />
          </div>
          <div>
            <label className="block text-sm font-medium text-gray-700 mb-1">
              Currency
            </label>
            <input
              type="text"
              name="currency"
              value={labelSettings.currency}
              onChange={handleCurrencyChange}
              className="mt-1 block w-full rounded-md border-gray-300 shadow-sm focus:border-blue-500 focus:ring-blue-500 sm:text-sm"
            />
          </div>
        </div>
      </div>

      <div className="flex justify-center">
        <button
          onClick={generatePDF}
          disabled={!data.length || isLoading}
          className={`px-4 py-2 rounded-md text-white font-medium ${!data.length || isLoading
            ? 'bg-gray-400 cursor-not-allowed'
            : 'bg-blue-600 hover:bg-blue-700'
            }`}
        >
          {isLoading ? 'Generating...' : 'Generate PDF Labels'}
        </button>
      </div>

      <div className="mt-8">
        <h2 className="text-lg font-semibold mb-2">Label Preview</h2>
        <p className='text-sm mb-4'>Not an accurate representation of the label</p>
        <div className="border border-gray-300 rounded-md p-4">
          <div
            style={{
              display: 'grid',
              gridTemplateColumns: `repeat(${labelSettings.columns}, 1fr)`,
              gap: `${labelSettings.rowGap}mm ${labelSettings.columnGap}mm`,
              padding: `${labelSettings.marginTop}mm ${labelSettings.marginLeft}mm`,
              width: `${labelSettings.pageWidth}mm`,
              maxHeight: `${labelSettings.pageHeight}mm`,
              overflow: 'hidden'
            }}
            className="bg-white border"
          >
            {preview.slice(0, labelSettings.labelsPerPage).map((item: any, idx: number) => (
              <div
                key={idx}
                style={{
                  width: `${labelSettings.labelWidth}mm`,
                  height: `${labelSettings.labelHeight}mm`,
                  minWidth: `${labelSettings.labelWidth}mm`,
                  minHeight: `${labelSettings.labelHeight}mm`
                }}
                className="bg-gray-100 p-1 rounded border border-gray-300 text-center flex flex-col justify-center overflow-hidden text-xs"
              >
                <div className="font-bold mb-1 text-sm leading-tight">
                  {formatCurrency(item.Price) || formatCurrency(item.price) || formatCurrency(item.PRICE) || '0.00'}
                </div>
                <div className="text-xs leading-tight mb-1 flex-1 overflow-hidden">
                  {item.Description || item.description || item.DESC || item.Name || item.name || 'Product Description'}
                </div>
                <div className="text-xs font-semibold">
                  {item.Code || item.code || item.ID || item.id || item.ItemNumber || `${idx + 1}`}
                </div>
              </div>
            ))}
          </div>
          {/* 
          <div className="grid grid-cols-2 gap-4">
            {preview.slice(0, 4).map((item: any, idx: number) => (
              <div key={idx} className="bg-gray-100 p-3 rounded border border-gray-300 text-center">
                <div className="text-lg font-bold">
                  {item.Price.toLocaleString(undefined, { maximumFractionDigits: 2, minimumFractionDigits: 2 }) || item.price.toLocaleString(undefined, { maximumFractionDigits: 2, minimumFractionDigits: 2 }) || item.PRICE.toLocaleString(undefined, { maximumFractionDigits: 2, minimumFractionDigits: 2 }) || '0.00'}
                </div>
                <div className="text-sm my-2 h-12 overflow-hidden">
                  {item.Description || item.description || item.DESC || item.Name || item.name || 'Product Description'}
                </div>
                <div className="text-base font-semibold">
                  {item.Code || item.code || item.ID || item.id || item.ItemNumber || `${idx + 1}`}
                </div>
              </div>
            ))}
          </div> */}
          {preview.length === 0 && (
            <div className="text-center py-8 text-gray-500">
              Upload an Excel file to see label previews
            </div>
          )}
        </div>
      </div>
    </div>
  );
}