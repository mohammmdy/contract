const express = require("express");
const ExcelJS = require("exceljs");
const topdf = require("docx2pdf-converter");
const multer = require("multer");
const path = require("path");
const cors = require("cors");
const fsN = require("fs/promises");
const os = require("os");
const dotenv = require("dotenv");



dotenv.config();


const app = express();

app.use(cors());
app.use(express.json());

const port = process.env.PORT || 3000;

const upload = multer({ dest: "uploads/" });


/**
 * Writes LOV values into a hidden sheet and returns a range reference.
 * This avoids Excel's 255-character limit on inline dropdown formulae.
 */
function buildLovRange(workbook, csvString, sheetName) {
  const values = csvString
    .split(",")
    .map((v) => v.trim())
    .filter(Boolean);

  // Use existing hidden sheet or create a new one
  let hiddenSheet = workbook.getWorksheet(sheetName);
  if (!hiddenSheet) {
    hiddenSheet = workbook.addWorksheet(sheetName);
    hiddenSheet.state = "veryHidden";
  }

  values.forEach((val, idx) => {
    hiddenSheet.getCell(`A${idx + 1}`).value = val;
  });

  return `${sheetName}!$A$1:$A$${values.length}`;
}

/**
 * Returns the Excel column letter for a given header name.
 * Looks up row 1 of the sheet to find the matching column.
 */
function getColByHeader(sheet, headerName) {
  const headerRow = sheet.getRow(1);
  let colLetter = null;

  headerRow.eachCell((cell, colNumber) => {
    if (cell.value === headerName) {
      colLetter = sheet.getColumn(colNumber).letter;
    }
  });

  if (!colLetter) {
    throw new Error(`Header "${headerName}" not found in sheet`);
  }

  return colLetter;
}

/**
 * Applies a dropdown to a column range, automatically choosing between
 */
function applyDropdown(workbook, sheet, headerName, csvString, sheetName) {
  const col = getColByHeader(sheet, headerName);
  const range = buildLovRange(workbook, csvString, sheetName);
  const formulae = [range];

  for (let i = 2; i <= 5000; i++) {
    sheet.getCell(`${col}${i}`).dataValidation = {
      type: "list",
      allowBlank: true,
      formulae,
    };
  }
}

async function createExcel(data) {
  const workbook = new ExcelJS.Workbook();
  const sheet = workbook.addWorksheet("Contracts");

  // Headers
  sheet.addRow(data.headerColumns);

  if (data.step == "1") {
    // Dropdown - Business Type
    applyDropdown(
      workbook,
      sheet,
      "Business Type",
      data.buLov,
      "LOV_BusinessType",
    );

    // Dropdown - Placeholder Type
    applyDropdown(
      workbook,
      sheet,
      "Placeholder Type",
      data.placeholdersLov,
      "LOV_Placeholder",
    );
  } else if (data.step == "2") {
    // Dropdown - Lead Negotiator
    applyDropdown(
      workbook,
      sheet,
      "Lead Negotiator",
      data.leadNegLov,
      "LOV_LeadNeg",
    );
    // Dropdown - Negotiation Status
    applyDropdown(
      workbook,
      sheet,
      "Negotiation Status",
      data.negotiationStatusLov,
      "LOV_NegotiationStatus",
    );
    // Dropdown - Contract Status
    applyDropdown(
      workbook,
      sheet,
      "Contract Status",
      data.contractStatusLov,
      "LOV_ContractStatus",
    );

    // Dropdown - Nahdi Signature Status
    applyDropdown(
      workbook,
      sheet,
      "Nahdi Signature Status",
      data.checkedLOV,
      "LOV_NahdiSignatureStatus",
    );

    // Dropdown - Vendor Signature Status
    applyDropdown(
      workbook,
      sheet,
      "Vendor Signature Status",
      data.checkedLOV,
      "LOV_VendorSignatureStatus",
    );
    // Dropdown - Is Purchase Commitment
    applyDropdown(
      workbook,
      sheet,
      "Is Purchase Commitment",
      data.checkedLOV,
      "LOV_IsPurchaseCommitment",
    );
    // Dropdown - Is Data Sharing
    applyDropdown(
      workbook,
      sheet,
      "Is Data Sharing",
      data.checkedLOV,
      "LOV_IsDataSharing",
    );
    // Dropdown - Renewable Contract Period
    applyDropdown(
      workbook,
      sheet,
      "Renewable Contract Period",
      data.checkedLOV,
      "LOV_RenewableContractPeriod",
    );
    // Dropdown - Authorization Letter Flag
    applyDropdown(
      workbook,
      sheet,
      "Authorization Letter Flag",
      data.checkedLOV,
      "LOV_AuthorizationLetterFlag",
    );
    // Dropdown - Is Returnable
    applyDropdown(
      workbook,
      sheet,
      "Is Returnable",
      data.checkedLOV,
      "LOV_IsReturnable",
    );
    // Dropdown - Is Folded Core
    applyDropdown(
      workbook,
      sheet,
      "Is Folded Core",
      data.checkedLOV,
      "LOV_IsFoldedCore",
    );
    // Dropdown - Is Folded Promo
    applyDropdown(
      workbook,
      sheet,
      "Is Folded Promo",
      data.checkedLOV,
      "LOV_IsFoldedPromo",
    );
    // Dropdown - Is Rebate
    applyDropdown(
      workbook,
      sheet,
      "Is Rebate",
      data.checkedLOV,
      "LOV_IsRebate",
    );
    // Dropdown - Is Autorenewal
    applyDropdown(
      workbook,
      sheet,
      "Is Autorenewal",
      data.checkedLOV,
      "LOV_IsAutorenewal",
    );
    // Dropdown - Requires Annex
    applyDropdown(
      workbook,
      sheet,
      "Requires Annex",
      data.checkedLOV,
      "LOV_RequiresAnnex",
    );
  } else if (data.step == "4") {
    // Dropdown - Business Unit
    applyDropdown(
      workbook,
      sheet,
      "Business Unit",
      data.buLov,
      "LOV_BusinessUnit",
    );

    // Dropdown - Is Advance Payment
    applyDropdown(
      workbook,
      sheet,
      "Is Advance Payment",
      data.checkedLOV,
      "LOV_AdvancePayment",
    );

    // Dropdown - Is Full Advance Payment
    applyDropdown(
      workbook,
      sheet,
      "Is Full Advance Payment",
      data.checkedLOV,
      "LOV_FullAdvancePayment",
    );
  } else if (data.step == "5") {
    if (data.tdType == "contract") {
      // Dropdown - Threshold Limit
      applyDropdown(
        workbook,
        sheet,
        "Threshold Limit",
        data.thresholdLov,
        "LOV_Threshold",
      );

      // Dropdown - Deal Comp Type
      applyDropdown(
        workbook,
        sheet,
        "Deal Comp Type Desc",
        data.dealCompLov,
        "LOV_DealComp",
      );

      // Dropdown - Rebate Sales
      applyDropdown(
        workbook,
        sheet,
        "Rebate Sales Ind",
        data.rebateSalesLov,
        "LOV_RebateSales",
      );

      // Dropdown - Restricted
      applyDropdown(
        workbook,
        sheet,
        "Restricted",
        data.restrictedLOV,
        "LOV_Restricted",
      );
    } else if (data.tdType == "products") {
      // Dropdown - Product List
      applyDropdown(
        workbook,
        sheet,
        "Product List",
        data.productListLOV,
        "LOV_ProductList",
      );

      // Dropdown - Flag
      applyDropdown(workbook, sheet, "Flag", data.flagLOV, "LOV_Flag");
    } else {
      // Dropdown - Product List
      applyDropdown(
        workbook,
        sheet,
        "Product List",
        data.productListLOV,
        "LOV_ProductList",
      );
    }
  } else if (data.step == "6") {
    // Dropdown - Element
    applyDropdown(workbook, sheet, "Element", data.annexJBPLov, "LOV_AnnexJBP");
  } else if (data.step == "7") {
    // Dropdown - Element
    applyDropdown(workbook, sheet, "Element", data.annexKBDLov, "LOV_AnnexKBD");
  }

  sheet.columns.forEach((col) => {
    col.width = 18;
  });

  return await workbook.xlsx.writeBuffer();
}

async function generateExcelTemplate(data, res) {
  const excelBuffer = await createExcel(data);

  res.setHeader("Content-Disposition", 'attachment; filename="Template.xlsx"');

  res.setHeader(
    "Content-Type",
    "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
  );

  res.send(excelBuffer);
}



app.post("/conServices/contractsExcelTemplate", async (req, res) => {
  await generateExcelTemplate(req.body, res);
});

app.post("/conServices/convertWordToPdf", upload.single("file"), async (req, res) => {
    const requestId = Date.now().toString();
    const requestFolder = path.join(__dirname, "uploads", requestId);

    try {
        if (!req.file) {
            return res.status(400).json({ error: "No file uploaded" });
        }
        await fsN.mkdir(requestFolder, { recursive: true });

        const inputPath = req.file.path;
        const originalName = req.file.originalname;
        // const pdfName = path.parse(originalName).name + ".pdf";
        const pdfName = "ContractTemplate.pdf";
        const pdfPath = path.join(requestFolder, pdfName);

        await topdf.convert(inputPath, pdfPath);

        const pdfBuffer = await fsN.readFile(pdfPath);
        const base64Pdf = pdfBuffer.toString("base64");

        res.on("finish", async () => {
            try {
                await fsN.rm(requestFolder, { recursive: true, force: true });
                await fsN.rm(inputPath, { force: true }); 
                console.log("Cleanup done");
            } catch (err) {
                console.error("Cleanup error:", err);
            }
        });

        res.json({
            fileName: pdfName,
            mimeType: "application/pdf",
            fileData: base64Pdf,
        });

    } catch (err) {
        // Clean up on error too
    try {
        await fsN.rm(requestFolder, { recursive: true, force: true });
        await fsN.rm(inputPath, { force: true });
    } catch {}
    
    console.error("Conversion error:", err);
    res.status(500).json({ error: "Conversion failed" });
    }
});


app.listen(port, () => {
  console.log(`Server running on http://localhost:${port}`);
});
