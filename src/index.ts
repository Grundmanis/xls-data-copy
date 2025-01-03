import express, { Request, Response, NextFunction } from "express";
import path from "path";
import fs from 'fs';
import multer from 'multer';
import xlsx, { WorkBook, WorkSheet } from 'xlsx';

const app = express();

app.use(express.static(path.join(__dirname, "../public")));

const uploadDir = path.join(__dirname, 'uploads');
if (!fs.existsSync(uploadDir)) {
  fs.mkdirSync(uploadDir, { recursive: true });
}

const storage = multer.diskStorage({
  destination: (req, file, cb) => {
    cb(null, 'uploads/'); // Define the upload folder
  },
  filename: (req, file, cb) => {
    cb(null, Date.now() + path.extname(file.originalname)); // Set a unique filename
  },
});
const upload = multer({ storage });


const packageTypes = [
  "AE - Aerosol",
  "AM - Ampoule, non - protected",
  "AP - Ampoule, protected",
  "AT - Atomizer",
  "BG - Bag",
  "FX - Bag, flexible container",
  "GY - Bag, gunny",
  "JB - Bag, jumbo",
  "ZB - Bag, large",
  "MB - Bag, multiply",
  "5M - Bag, paper",
  "XJ - Bag, paper, multi - wall",
  "XK - Bag, paper, multi - wall, water resistant",
  "EC - Bag, plastic",
  "XD - Bag, plastics film",
  "44 - Bag, polybag",
  "43 - Bag, super bulk",
  "5L - Bag, textile",
  "XG - Bag, textile, sift proof",
  "XH - Bag, textile, water resistant",
  "XF - Bag, textile, without inner coat/liner",
  "TT - Bag, tote",
  "5H - Bag, woven plastic",
  "XB - Bag, woven plastic, sift proof",
  "XC - Bag, woven plastic, water resistant",
  "XA - Bag, woven plastic, without inner coat/liner",
  "BL - Bale, compressed",
  "BN - Bale, non - compressed",
  "AL - Ball",
  "BF - Balloon, non - protected",
  "BP - Balloon, protected",
  "BR - Bar",
  "BA - Barrel",
  "2C - Barrel, wooden",
  "QH - Barrel, wooden, bung type",
  "QJ - Barrel, wooden, removable head",
  "BZ - Bars, in bundle/bunch/truss",
  "BM - Basin",
  "BK - Basket",
  "HC - Basket, with handle, cardboard",
  "HA - Basket, with handle, plastic",
  "HB - Basket, with handle, wooden",
  "B4 - Belt",
  "BI - Bin",
  "OK - Block",
  "BD - Board",
  "BY - Board, in bundle/bunch/truss",
  "BB - Bobbin",
  "BT - Bolt",
  "GB - Bottle, gas",
  "BS - Bottle, non - protected, bulbous",
  "BO - Bottle, non - protected, cylindrical",
  "BV - Bottle, protected bulbous",
  "BQ - Bottle, protected cylindrical",
  "BC - Bottlecrate / bottlerack",
  "BX - Box",
  "4B - Box, aluminium",
  "DH - Box, Commonwealth Handling Equipment Pool (CHEP), Eurobox",
  "4G - Box, fibreboard",
  "BW - Box, for liquids",
  "4C - Box, natural wood",
  "4H - Box, plastic",
  "QR - Box, plastic, expanded",
  "QS - Box, plastic, solid",
  "4D - Box, plywood",
  "4F - Box, reconstituted wood",
  "4A - Box, steel",
  "QP - Box, wooden, natural wood, ordinary",
  "QQ - Box, wooden, natural wood, with sift proof walls",
  "BJ - Bucket",
  "VG - Bulk, gas (at 1031 mbar and 15 degree C)",
  "VQ - Bulk, liquefied gas (at abnormal temperature/pressure)",
  "VL - Bulk, liquid",
  "VS - Bulk, scrap metal",
  "VY - Bulk, solid, fine particles (\"powders\")",
  "VR - Bulk, solid, granular particles (\"grains\")",
  "VO - Bulk, solid, large particles (\"nodules\")",
  "BH - Bunch",
  "BE - Bundle",
  "8C - Bundle, wooden",
  "BU - Butt",
  "CG - Cage",
  "DG - Cage, Commonwealth Handling Equipment Pool (CHEP)",
  "CW - Cage, roll",
  "CX - Can, cylindrical",
  "CA - Can, rectangular",
  "CD - Can, with handle and spout",
  "CI - Canister",
  "CZ - Canvas",
  "AV - Capsule",
  "CO - Carboy, non - protected",
  "CP - Carboy, protected",
  "CM - Card",
  "FW - Cart, flatbed",
  "CT - Carton",
  "CQ - Cartridge",
  "CS - Case",
  "7A - Case, car",
  "EI - Case, isothermic",
  "SK - Case, skeleton",
  "SS - Case, steel",
  "ED - Case, with pallet base",
  "EF - Case, with pallet base, cardboard",
  "EH - Case, with pallet base, metal",
  "EG - Case, with pallet base, plastic",
  "EE - Case, with pallet base, wooden",
  "7B - Case, wooden",
  "CK - Cask",
  "CH - Chest",
  "CC - Churn",
  "AI - Clamshell",
  "CF - Coffer",
  "CJ - Coffin",
  "CL - Coil",
  "6P - Composite packaging, glass receptacle",
  "YR - Composite packaging, glass receptacle in aluminium crate",
  "YQ - Composite packaging, glass receptacle in aluminium drum",
  "YY - Composite packaging, glass receptacle in expandable plastic pack",
  "YW - Composite packaging, glass receptacle in fibre drum",
  "YX - Composite packaging, glass receptacle in fibreboard box",
  "YT - Composite packaging, glass receptacle in plywood drum",
  "YZ - Composite packaging, glass receptacle in solid plastic pack",
  "YP - Composite packaging, glass receptacle in steel crate box",
  "YN - Composite packaging, glass receptacle in steel drum",
  "YV - Composite packaging, glass receptacle in wickerwork hamper",
  "YS - Composite packaging, glass receptacle in wooden box",
  "6H - Composite packaging, plastic receptacle",
  "YD - Composite packaging, plastic receptacle in aluminium crate",
  "YC - Composite packaging, plastic receptacle in aluminium drum",
  "YJ - Composite packaging, plastic receptacle in fibre drum",
  "YK - Composite packaging, plastic receptacle in fibreboard box",
  "YL - Composite packaging, plastic receptacle in plastic drum",
  "YH - Composite packaging, plastic receptacle in plywood box",
  "YG - Composite packaging, plastic receptacle in plywood drum",
  "YM - Composite packaging, plastic receptacle in solid plastic box",
  "YB - Composite packaging, plastic receptacle in steel crate box",
  "YA - Composite packaging, plastic receptacle in steel drum",
  "YF - Composite packaging, plastic receptacle in wooden box",
  "AJ - Cone",
  "1F - Container, flexible",
  "GL - Container, gallon",
  "ME - Container, metal",
  "CN - Container, not otherwise specified as transport equipment",
  "OU - Container, outer",
  "CV - Cover",
  "CR - Crate",
  "CB - Crate, beer",
  "DK - Crate, bulk, cardboard",
  "DL - Crate, bulk, plastic",
  "DM - Crate, bulk, wooden",
  "FD - Crate, framed",
  "FC - Crate, fruit",
  "MA - Crate, metal",
  "MC - Crate, milk",
  "DC - Crate, multiple layer, cardboard",
  "DA - Crate, multiple layer, plastic",
  "DB - Crate, multiple layer, wooden",
  "SC - Crate, shallow",
  "8B - Crate, wooden",
  "CE - Creel",
  "CU - Cup",
  "CY - Cylinder",
  "DJ - Demijohn, non - protected",
  "DP - Demijohn, protected",
  "DN - Dispenser",
  "DR - Drum",
  "1B - Drum, aluminium",
  "QC - Drum, aluminium, non - removable head",
  "QD - Drum, aluminium, removable head",
  "1G - Drum, fibre",
  "DI - Drum, iron",
  "IH - Drum, plastic",
  "QF - Drum, plastic, non - removable head",
  "QG - Drum, plastic, removable head",
  "1D - Drum, plywood",
  "1A - Drum, steel",
  "QA - Drum, steel, non - removable head",
  "QB - Drum, steel, removable head",
  "1W - Drum, wooden",
  "EN - Envelope",
  "SV - Envelope, steel",
  "FP - Filmpack",
  "FI - Firkin",
  "FL - Flask",
  "FB - Flexibag",
  "FE - Flexitank",
  "FT - Foodtainer",
  "FO - Footlocker",
  "FR - Frame",
  "GI - Girder",
  "GZ - Girders, in bundle/bunch/truss",
  "HR - Hamper",
  "HN - Hanger",
  "HG - Hogshead",
  "IN - Ingot",
  "IZ - Ingots, in bundle/bunch/truss",
  "WA - Intermediate bulk container",
  "WD - Intermediate bulk container, aluminium",
  "WL - Intermediate bulk container, aluminium, liquid",
  "WH - Intermediate bulk container, aluminium, pressurised &gt; 10 kpa",
  "ZS - Intermediate bulk container, composite",
  "ZR - Intermediate bulk container, composite, flexible plastic, liquids",
  "ZP - Intermediate bulk container, composite, flexible plastic, pressurised",
  "ZM - Intermediate bulk container, composite, flexible plastic, solids",
  "ZQ - Intermediate bulk container, composite, rigid plastic, liquids",
  "ZN - Intermediate bulk container, composite, rigid plastic, pressurised",
  "ZL - Intermediate bulk container, composite, rigid plastic, solids",
  "ZT - Intermediate bulk container, fibreboard",
  "ZU - Intermediate bulk container, flexible",
  "WF - Intermediate bulk container, metal",
  "WM - Intermediate bulk container, metal, liquid",
  "ZV - Intermediate bulk container, metal, other than steel",
  "WJ - Intermediate bulk container, metal, pressure 10 kpa",
  "ZW - Intermediate bulk container, natural wood",
  "WU - Intermediate bulk container, natural wood, with inner liner",
  "ZA - Intermediate bulk container, paper, multi - wall",
  "ZC - Intermediate bulk container, paper, multi - wall, water resistant",
  "WS - Intermediate bulk container, plastic film",
  "ZX - Intermediate bulk container, plywood",
  "WY - Intermediate bulk container, plywood, with inner liner",
  "ZY - Intermediate bulk container, reconstituted wood",
  "WZ - Intermediate bulk container, reconstituted wood, with inner liner",
  "AA - Intermediate bulk container, rigid plastic",
  "ZK - Intermediate bulk container, rigid plastic, freestanding, liquids",
  "ZH - Intermediate bulk container, rigid plastic, freestanding, pressurised",
  "ZF - Intermediate bulk container, rigid plastic, freestanding, solids",
  "ZJ - Intermediate bulk container, rigid plastic, with structural equipment, liquids",
  "ZG - Intermediate bulk container, rigid plastic, with structural equipment, pressurised",
  "ZD - Intermediate bulk container, rigid plastic, with structural equipment, solids",
  "WC - Intermediate bulk container, steel",
  "WK - Intermediate bulk container, steel, liquid",
  "WG - Intermediate bulk container, steel, pressurised &gt; 10 kpa",
  "WT - Intermediate bulk container, textile with out coat/liner",
  "WV - Intermediate bulk container, textile, coated",
  "WX - Intermediate bulk container, textile, coated and liner",
  "WW - Intermediate bulk container, textile, with liner",
  "WP - Intermediate bulk container, woven plastic, coated",
  "WR - Intermediate bulk container, woven plastic, coated and liner",
  "WQ - Intermediate bulk container, woven plastic, with liner",
  "WN - Intermediate bulk container, woven plastic, without coat/liner",
  "JR - Jar",
  "JY - Jerrican, cylindrical",
  "3H - Jerrican, plastic",
  "QM - Jerrican, plastic, non - removable head",
  "QN - Jerrican, plastic, removable head",
  "JC - Jerrican, rectangular",
  "3A - Jerrican, steel",
  "QK - Jerrican, steel, non - removable head",
  "QL - Jerrican, steel, removable head",
  "JG - Jug",
  "JT - Jutebag",
  "KG - Keg",
  "KI - Kit",
  "LV - Liftvan",
  "LG - Log",
  "LZ - Logs, in bundle/bunch/truss",
  "LT - Lot",
  "LU - Lug",
  "LE - Luggage",
  "MT - Mat",
  "MX - Matchbox",
  "ZZ - Mutually defined",
  "NS - Nest",
  "NT - Net",
  "NU - Net, tube, plastic",
  "NV - Net, tube, textile",
  "NA - Not available",
  "OT - Octabin",
  "PK - Package",
  "IK - Package, cardboard, with bottle grip - holes",
  "IB - Package, display, cardboard",
  "ID - Package, display, metal",
  "IC - Package, display, plastic",
  "IA - Package, display, wooden",
  "IF - Package, flow",
  "IG - Package, paper wrapped",
  "IE - Package, show",
  "PA - Packet",
  "PL - Pail",
  "PX - Pallet",
  "AH - Pallet, 100cms × 110cms",
  "OD - Pallet, AS 4068 - 1993",
  "PB - Pallet, box",
  "OC - Pallet, CHEP 100 cm x 120 cm",
  "OA - Pallet, CHEP 40 cm x 60 cm",
  "OB - Pallet, CHEP 80 cm x 120 cm",
  "OE - Pallet, ISO T11",
  "PD - Pallet, modular, collars 80cms * 100cms",
  "PE - Pallet, modular, collars 80cms * 120cms",
  "AF - Pallet, modular, collars 80cms × 60cms",
  "AG - Pallet, shrinkwrapped",
  "TW - Pallet, triwall",
  "8A - Pallet, wooden",
  "P2 - Pan",
  "PC - Parcel",
  "PF - Pen",
  "PP - Piece",
  "PI - Pipe",
  "PV - Pipes, in bundle/bunch/truss",
  "PH - Pitcher",
  "PN - Plank",
  "PZ - Planks, in bundle/bunch/truss",
  "PG - Plate",
  "PY - Plates, in bundle/bunch/truss",
  "OF - Platform, unspecified weight or dimension",
  "PT - Pot",
  "PO - Pouch",
  "PJ - Punnet",
  "RK - Rack",
  "RJ - Rack, clothing hanger",
  "AB - Receptacle, fibre",
  "GR - Receptacle, glass",
  "MR - Receptacle, metal",
  "AC - Receptacle, paper",
  "PR - Receptacle, plastic",
  "MW - Receptacle, plastic wrapped",
  "AD - Receptacle, wooden",
  "RT - Rednet",
  "RL - Reel",
  "RG - Ring",
  "RD - Rod",
  "RZ - Rods, in bundle/bunch/truss",
  "RO - Roll",
  "SH - Sachet",
  "SA - Sack",
  "MS - Sack, multi - wall",
  "SE - Sea - chest",
  "ST - Sheet",
  "SP - Sheet, plastic wrapping",
  "SM - Sheetmetal",
  "SZ - Sheets, in bundle/bunch/truss",
  "SW - Shrinkwrapped",
  "SI - Skid",
  "SB - Slab",
  "SY - Sleeve",
  "SL - Slipsheet",
  "SD - Spindle",
  "SO - Spool",
  "SU - Suitcase",
  "T1 - Tablet",
  "TG - Tank container, generic",
  "TY - Tank, cylindrical",
  "TK - Tank, rectangular",
  "TC - Tea - chest",
  "TI - Tierce TI",
  "TN - Tin",
  "PU - Tray pack",
  "GU - Tray, containing horizontally stacked flat items",
  "DV - Tray, one layer no cover, cardboard",
  "DS - Tray, one layer no cover, plastic",
  "DU - Tray, one layer no cover, polystyrene",
  "DT - Tray, one layer no cover, wooden",
  "IL - Tray, rigid, lidded stackable (CEN TS 14482:2002)",
  "DY - Tray, two layers no cover, cardboard",
  "DW - Tray, two layers no cover, plastic tray",
  "DX - Tray, two layers no cover, wooden",
  "TR - Trunk",
  "TS - Truss",
  "TB - Tub",
  "TL - Tub, with lid",
  "TU - Tube",
  "TD - Tube, collapsible",
  "TV - Tube, with nozzle",
  "TZ - Tubes, in bundle/bunch/truss",
  "TO - Tun",
  "TE - Tyre",
  "UC - Uncaged",
  "UN - Unit",
  "NE - Unpacked or unpackaged",
  "NG - Unpacked or unpackaged, multiple units",
  "NF - Unpacked or unpackaged, single unit",
  "VP - Vacuum - packed",
  "VK - Vanpack",
  "VA - Vat",
  "VN - Vehicle",
  "VI - Vial",
  "WB - Wickerbottle",
]


// Endpoint to upload two xlsx files
  // @ts-ignore
app.post('/upload', upload.fields([{ name: 'from' }, { name: 'to' }]), (req, res) => {
  const { from, to } = req.files as { from: Express.Multer.File[], to: Express.Multer.File[] };

  if (!from || !to) {
    return res.status(400).send('Both files are required.');
  }

  // Read the first file
  const workbook1 = xlsx.readFile(from[0].path);
  const sourceSheet = workbook1.Sheets[workbook1.SheetNames[0]]; // Assuming the first sheet

  // Read the second file
  const workbook2 = xlsx.readFile(to[0].path);

  putData("Dangerous_Cargo", workbook2, sourceSheet, res, 5, 1);
  putData("Cargo", workbook2, sourceSheet, res, 5, 1);
  putData("Cargo_partners", workbook2, sourceSheet, res, 6, 2);

  // Save the updated target file
  const outputFile = path.join('uploads', 'updated_' + to[0].filename);
  xlsx.writeFile(workbook2, outputFile);

  // Send the file back for download
  res.download(outputFile, 'updated_' + to[0].filename, (err) => {
    if (err) {
      console.error('Error sending file:', err);
      return res.status(500).send('Error sending file');
    }
    console.log('File sent successfully');
  });
});

const putData = (targetTab: string, workbook2: WorkBook, sourceSheet: WorkSheet, res: Response, startRow: number, targetRow: number) => {
  // Create an array to store all updates for the target sheet
  const updates: { [key: string]: any } = {}; // Key: cell address, Value: cell value
  const targetSheetIndex = workbook2.SheetNames.indexOf(targetTab);
  if (targetSheetIndex === -1) {
    return res.status(400).send('Target tab not found.');
  }

  const targetSheet = workbook2.Sheets[workbook2.SheetNames[targetSheetIndex]]; // Assuming the first sheet
  if (!targetSheet) {
    console.error('Target sheet not found');
    return;
  }

  // Get the 2nd row from the target sheet (column mapping)
  const targetRow2 = xlsx.utils.sheet_to_json(targetSheet, { header: 1 })[targetRow]; // Row 2 (index 1)
  
  // Get the data from the source sheet starting from row 3 (index 2)
  const sourceData = xlsx.utils.sheet_to_json(sourceSheet, { header: 1, range: 2 }); // Skip the first 2 rows

  // Loop through each column in the target sheet's second row (targetRow2) to map the columns
  // @ts-expect-error
  
// Loop through each column in the target sheet's second row (targetRow2) to map the columns
  targetRow2.forEach((columnMapping, colIndex) => {
    if (columnMapping) {
      console.log("columnMapping",columnMapping);
      const includesAmpersant = columnMapping.includes('&');
      const columns = columnMapping.includes('/') || columnMapping.includes('&') 
      ? columnMapping.replace(/[\/&\s]/g, '').split('') 
      : columnMapping.length > 2 
        ? columnMapping.split('') 
        : columnMapping.match(/.{1}/g); // Handle cases like "A/B", "AO & AQ", "AO&AQ", and "AB"

    console.log("columns", columns);
  
      // Iterate over each row in the source data
      sourceData.forEach((sourceRow, rowIndex) => {
        // @ts-ignore
        const sourceValues = columns.map((col) => {
          const colLetter = col.trim();
          const sourceColIndex = xlsx.utils.decode_col(colLetter); // Get column index from letter
          // @ts-ignore
          return sourceRow[sourceColIndex]; // Extract value for the current row
        });

        console.log("sourceValues",sourceValues);

        // if (sourceValues && sourceValues.length > 0) {
          // @ts-ignore
          // sourceValues.forEach((value, valueIndex) => {
            // if (!value) {
            //   return; // Skip empty values
            // }
            let finalValue = sourceValues.length > 1 ? sourceValues.join(' ') : sourceValues[0];
  
            // Check if the value starts with any element from the array
            packageTypes.forEach((packageType) => {
              if (finalValue && packageType.startsWith(finalValue)) {
                  finalValue = packageType; // Replace with the matching full value
              }
            });
            console.log("finalValue", finalValue);
  
            const targetCellAddress = xlsx.utils.encode_cell({
              r: startRow + rowIndex,
              c: colIndex,
            });

            // Get the existing cell object to preserve styles
            const existingCell = targetSheet[targetCellAddress] || {};
  
            // Preserve the style and update the value
            updates[targetCellAddress] = {
              ...existingCell, // Preserve the existing cell properties (styles, etc.)
              v: finalValue, // Set the mapped or original value
            };
          // });
        // }
      });
    }
  });

  console.log({updates});

  // Apply all updates to the target sheet in one go
  Object.keys(updates).forEach((cell) => {
      targetSheet[cell] = updates[cell]; // Apply each update to the target sheet
    });

  // console.log("targetSheet", targetSheet);
  const currentRef = targetSheet['!ref'];
  const newRef = 'A1:Y500'; // Dynamically determine based on your needs
  if (currentRef !== newRef) {
      targetSheet['!ref'] = newRef;
  }
}

app.get("/", (req: Request, res: Response, next: NextFunction): void => {
  try {
    console.log("index");
    res.send("index.html");
  } catch (error) {
    next(error);
  }
});

const PORT = 3000;

app.listen(PORT, () => {
  console.log(`App listening on port ${PORT}`);
});