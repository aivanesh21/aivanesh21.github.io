/*!
 * GSTR-1 Converter engine — Sawant & Kadrekar LLP
 * Converts a Munim "GSTR-1 Detail" sales / CDN export into
 *   (a) the GSTN GSTR-1 Excel Workbook Template (V2.1) used by the offline tool, and
 *   (b) a GSTR-1 JSON for direct upload on the GST portal.
 * Pure JavaScript, no DOM. Runs in a Web Worker in the browser and in Node for tests.
 * Depends on fflate (MIT) for zip I/O: global `fflate` or require('fflate').
 */
(function (root) {
  'use strict';

  var fflate = root.fflate || (typeof globalThis !== 'undefined' && globalThis.fflate)
    || (typeof require === 'function' ? require('fflate') : null);

  /* ------------------------------------------------------------------ *
   * 1. Reference data
   * ------------------------------------------------------------------ */

  var TEMPLATE = {"sheets":[{"n":"b2b,sez,de","r1":["Summary For B2B, SEZ, DE (4A, 4B, 6B, 6C)",null,null,null,null,null,null,null,null,null,null,null,"HELP"],"r2":["No. of Recipients",null,"No. of Invoices",null,"Total Invoice Value",null,null,null,null,null,null,"Total Taxable Value","Total Cess"],"h":["GSTIN/UIN of Recipient","Receiver Name","Invoice Number","Invoice date","Invoice Value","Place Of Supply","Reverse Charge","Applicable % of Tax Rate","Invoice Type","E-Commerce GSTIN","Rate","Taxable Value","Cess Amount"],"w":[22.7,22.7,17.8,11.3,19.5,20.0,14.8,14.8,13.5,19.3,7.5,27.5,22.2]},{"n":"b2ba","r1":["Summary For B2BA","Original details ",null,null,"Revised Details ",null,null,null,null,null,null,null,null,null,"HELP"],"r2":["No. of Recipients",null,"No. of Invoices",null,null,null,"Total Invoice Value",null,null,null,null,null,null,"Total Taxable Value","Total Cess"],"h":["GSTIN/UIN of Recipient","Receiver Name","Original Invoice Number","Original Invoice date","Revised Invoice Number","Revised Invoice date","Invoice Value","Place Of Supply","Reverse Charge","Applicable % of Tax Rate","Invoice Type","E-Commerce GSTIN","Rate","Taxable Value","Cess Amount"],"w":[25.5,13.0,29.7,26.3,24.0,22.0,19.2,20.2,16.8,14.8,15.2,19.5,12.2,21.3,16.2]},{"n":"b2cl","r1":["Summary For B2CL(5)",null,null,null,null,null,null,null,"HELP"],"r2":["No. of Invoices"],"h":["Invoice Number","Invoice date","Invoice Value","Place Of Supply","Applicable % of Tax Rate","Rate","Taxable Value","Cess Amount","E-Commerce GSTIN"],"w":[22.7,11.3,19.5,19.5,14.8,5.7,20.0,12.5,19.3]},{"n":"b2cla","r1":["Summary For B2CLA","Original details ",null,null,"Revised Details ",null,null,null,null,null,"HELP"],"r2":["No. of Invoices",null,null,null,null,"Total Inv Value",null,null,"Total Taxable Value","Total Cess"],"h":["Original Invoice Number","Original Invoice date","Original Place Of Supply","Revised Invoice Number","Revised Invoice date","Invoice Value","Applicable % of Tax Rate","Rate","Taxable Value","Cess Amount","E-Commerce GSTIN"],"w":[24.5,18.5,24.5,29.7,26.3,17.8,14.8,7.0,19.0,12.5,19.7]},{"n":"b2cs","r1":["Summary For B2CS(7)",null,null,null,null,null,"HELP"],"r2":[null,null,null,null,"Total Taxable  Value","Total Cess"],"h":["Type","Place Of Supply","Applicable % of Tax Rate","Rate","Taxable Value","Cess Amount","E-Commerce GSTIN"],"w":[22.5,19.5,14.8,5.5,20.5,12.5,19.3]},{"n":"b2csa","r1":["Summary For B2CSA","Original details ","Revised details",null,null,null,null,null,"HELP"],"r2":[null,null,null,null,null,null,"Total Taxable  Value","Total Cess"],"h":["Financial Year","Original Month","Place Of Supply","Type","Applicable % of Tax Rate","Rate","Taxable Value","Cess Amount","E-Commerce GSTIN"],"w":[22.5,13.0,13.0,29.8,16.0,14.8,16.0,21.2,19.3]},{"n":"cdnr","r1":["Summary For CDNR(9B)",null,null,null,null,null,null,null,null,null,null,null,"HELP"],"r2":["No. of Recipients",null,"No. of Notes",null,null,null,null,null,"Total Note Value",null,null,"Total Taxable Value","Total Cess"],"h":["GSTIN/UIN of Recipient","Receiver Name","Note Number","Note Date","Note Type","Place Of Supply","Reverse Charge","Note Supply Type","Note Value","Applicable % of Tax Rate","Rate","Taxable Value","Cess Amount"],"w":[24.8,24.8,27.7,24.8,14.7,19.5,14.8,26.7,20.8,14.8,6.3,20.0,12.5]},{"n":"cdnra","r1":["Summary For CDNRA","Original details ",null,null,"Revised details",null,null,null,null,null,null,null,null,null,"HELP"],"r2":["No. of Recipients",null,"No. of Notes/Vouchers",null,null,null,null,null,null,null,"Total Note Value",null,null,"Total Taxable Value","Total Cess"],"h":["GSTIN/UIN of Recipient","Receiver Name","Original Note Number","Original Note Date","Revised Note Number","Revised Note Date","Note Type","Place Of Supply","Reverse Charge","Note Supply Type","Note Value","Applicable % of Tax Rate","Rate","Taxable Value","Cess Amount"],"w":[26.7,13.0,34.5,31.5,34.7,30.3,14.8,19.5,14.8,26.7,34.5,14.8,8.2,19.2,15.3]},{"n":"cdnur","r1":["Summary For CDNUR(9B)",null,null,null,null,null,null,null,null,"HELP"],"r2":[null,"No. of Notes/Vouchers",null,null,null,"Total Note Value",null,null,"Total Taxable Value","Total Cess"],"h":["UR Type","Note Number","Note Date","Note Type","Place Of Supply","Note Value","Applicable % of Tax Rate","Rate","Taxable Value","Cess Amount"],"w":[26.5,27.7,25.5,14.7,19.5,25.2,14.8,5.5,20.0,12.5]},{"n":"cdnura","r1":["Summary For CDNURA","Original details ",null,"Revised details",null,null,null,null,null,null,null,"HELP"],"r2":[null,"No. of Notes/Vouchers",null,null,null,null,null,"Total Note Value",null,null,"Total Taxable Value","Total Cess"],"h":["UR Type","Original Note Number","Original Note Date","Revised Note Number","Revised Note Date","Note Type","Place Of Supply","Note Value","Applicable % of Tax Rate","Rate","Taxable Value","Cess Amount"],"w":[23.5,35.5,29.5,36.2,33.3,18.3,19.5,27.5,14.8,11.0,20.7,14.5]},{"n":"exp","r1":["Summary For EXP(6)",null,null,null,null,null,null,null,null,"HELP"],"r2":[null,"No. of Invoices",null,"Total Invoice Value",null,"No. of Shipping Bill",null,null,null,"Total Taxable Value"],"h":["Export Type","Invoice Number","Invoice date","Invoice Value","Port Code","Shipping Bill Number","Shipping Bill Date","Rate","Taxable Value","Cess Amount"],"w":[21.0,15.3,11.3,19.5,9.5,19.3,16.2,5.5,20.0,13.0]},{"n":"expa","r1":["Summary For EXPA","Original details ",null,"Revised details",null,null,null,null,null,null,null,"HELP"],"r2":[null,"No. of Invoices",null,null,null,"Total Invoice Value",null,"No. of Shipping Bill",null,null,"Total Taxable Value","Total Cess"],"h":["Export Type","Original Invoice Number","Original Invoice date","Revised Invoice Number","Revised Invoice date","Invoice Value","Port Code","Shipping Bill Number","Shipping Bill Date","Rate","Taxable Value","Cess Amount"],"w":[23.0,22.8,20.0,29.7,26.3,19.8,15.0,23.2,16.2,9.3,19.5,20.0]},{"n":"at","r1":["Summary For Advance Received (11B) ",null,null,null,"HELP"],"r2":[null,null,null,"Total Advance Received","Total Cess"],"h":["Place Of Supply","Applicable % of Tax Rate","Rate","Gross Advance Received","Cess Amount"],"w":[28.5,14.8,5.5,24.0,12.5]},{"n":"ata","r1":["Summary For Amended Tax Liability(Advance Received) ","Original details ",null,"Revised details",null,null,"HELP"],"r2":[null,null,null,null,null,"Total Advance Received","Total Cess"],"h":["Financial Year","Original Month","Original Place Of Supply","Applicable % of Tax Rate","Rate","Gross Advance Received","Cess Amount"],"w":[26.7,26.7,21.7,14.8,12.5,25.8,15.3]},{"n":"atadj","r1":["Summary For Advance Adjusted (11B) ",null,null,null,"HELP"],"r2":[null,null,null,"Total Advance Adjusted","Total Cess"],"h":["Place Of Supply","Applicable % of Tax Rate","Rate","Gross Advance Adjusted","Cess Amount"],"w":[27.8,14.8,5.5,24.0,12.5]},{"n":"atadja","r1":["Summary For Amendement Of Adjustment Advances","Original details ",null,"Revised details",null,null,"HELP"],"r2":[null,null,null,null,null,"Total Advance Adjusted","Total Cess"],"h":["Financial Year","Original Month","Original Place Of Supply","Applicable % of Tax Rate","Rate","Gross Advance Adjusted","Cess Amount"],"w":[26.7,26.7,23.5,14.8,9.2,24.0,12.5]},{"n":"exemp","r1":["Summary For Nil rated, exempted and non GST outward supplies (8)",null,null,"HELP"],"r2":[null,"Total Nil Rated Supplies","Total Exempted Supplies","Total Non-GST Supplies"],"h":["Description","Nil Rated Supplies","Exempted(other than nil rated/non GST supply)","Non-GST Supplies"],"w":[38.2,24.2,24.8,24.0]},{"n":"hsn(b2b)","r1":["Summary For HSN(12)",null,null,null,null,null,null,null,null,null,"HELP"],"r2":["No. of HSN",null,null,null,"Total Value",null,"Total Taxable Value","Total Integrated Tax","Total Central Tax","Total State/UT Tax","Total Cess"],"h":["HSN","Description","UQC","Total Quantity","Total Value","Rate","Taxable Value","Integrated Tax Amount","Central Tax Amount","State/UT Tax Amount","Cess Amount"],"w":[22.5,14.2,19.3,13.5,23.7,15.3,29.2,21.0,18.5,20.2,12.5]},{"n":"hsn(b2c)","r1":["Summary For HSN(12)",null,null,null,null,null,null,null,null,null,"HELP"],"r2":["No. of HSN",null,null,null,"Total Value",null,"Total Taxable Value","Total Integrated Tax","Total Central Tax","Total State/UT Tax","Total Cess"],"h":["HSN","Description","UQC","Total Quantity","Total Value","Rate","Taxable Value","Integrated Tax Amount","Central Tax Amount","State/UT Tax Amount","Cess Amount"],"w":[22.5,14.2,19.3,13.5,23.7,15.3,29.2,21.0,18.5,20.2,12.5]},{"n":"docs","r1":["Summary of documents issued during the tax period (13)",null,null,null,"HELP"],"r2":[null,null,null,"Total Number","Total Cancelled"],"h":["Nature of Document","Sr. No. From","Sr. No. To","Total Number","Cancelled"],"w":[50.7,12.2,11.0,14.2,15.8]},{"n":"eco","r1":["Summary For Supplies through ECO-14",null,null,null,null,null,null,"HELP"],"r2":[null,"No. of E-Commerce Operator",null,"Total Net Value of Supplies","Total Integrated Tax","Total Central Tax ","Total State/UT Tax ","Total Cess"],"h":["Nature of Supply","GSTIN of E-Commerce Operator","E-Commerce Operator Name","Net value of supplies","Integrated tax","Central tax","State/UT tax","Cess"],"w":[42.8,33.5,29.2,23.7,25.5,23.0,23.7,23.2]},{"n":"ecoa","r1":["Summary For Amended Supplies through ECO-14A","Original details ",null,"Revised details",null,null,null,null,null,null,"HELP"],"r2":[null,null,null,"No. of E-Commerce Operator",null,null,"Total Net Value of Supplies","Total Integrated Tax","Total Central Tax ","Total State/UT Tax ","Total Cess"],"h":["Nature of Supply","Financial Year","Original Month/Quarter","Original GSTIN of E-Commerce Operator","Revised GSTIN of E-Commerce Operator","E-Commerce Operator Name","Revised Net value of supplies","Integrated tax","Central tax","State/UT tax","Cess"],"w":[55.3,14.5,23.7,42.5,38.5,29.2,29.5,26.8,23.5,24.0,21.3]},{"n":"ecob2b","r1":["Summary For Supplies U/s 9(5)-15-B2B",null,null,null,null,null,null,null,null,null,null,"HELP"],"r2":["No. of Supplier",null,"No. of Recipients",null,"No. of Documents",null,"Total value of supplies made",null,null,null,"Total Taxable Value","Total Cess"],"h":["Supplier GSTIN/UIN","Supplier Name","Recipient GSTIN/UIN","Recipient Name","Document Number","Document Date","Value of supplies made","Place Of Supply","Document type","Rate","Taxable Value","Cess Amount"],"w":[41.5,25.8,25.0,26.3,21.5,15.5,30.0,25.8,22.8,12.2,27.5,25.8]},{"n":"ecourp2b","r1":["Summary For Supplies U/s 9(5)-15-URP2B",null,null,null,null,null,null,null,null,"HELP"],"r2":["No. of Recipients",null,"No. of Documents",null,"Total value of supplies made",null,null,null,"Total Taxable Value","Total Cess"],"h":["Recipient GSTIN/UIN","Recipient Name","Document Number","Document Date","Value of supplies made","Place Of Supply","Document type","Rate","Taxable Value","Cess Amount"],"w":[44.5,21.0,19.2,15.5,30.3,23.3,18.8,10.5,27.2,27.5]},{"n":"ecob2c","r1":["Summary For Supplies U/s 9(5)-15-B2C",null,null,null,null,"HELP"],"r2":["No. of Supplier",null,null,"Total Taxable Value",null,"Total Cess"],"h":["Supplier GSTIN/UIN","Supplier Name","Place Of Supply","Taxable Value","Rate","Cess Amount"],"w":[41.7,26.2,27.5,25.8,14.3,34.5]},{"n":"ecourp2c","r1":["Summary For Supplies U/s 9(5)-15-URP2C",null,null,"HELP"],"r2":[null,"Total Taxable Value",null,"Total Cess"],"h":["Place Of Supply","Taxable Value","Rate","Cess Amount"],"w":[44.8,24.0,15.0,24.3]},{"n":"ecoab2b","r1":["Summary For Supplies U/s 9(5) - 15A-B2B","Original details ",null,null,null,null,"Revised details ",null,null,null,null,null,null,"HELP"],"r2":["No. of Supplier",null,"No. of Recipients",null,"No. of Documents",null,null,null,"Total value of supplies made",null,null,null,"Total Taxable Value","Total Cess"],"h":["Supplier GSTIN/UIN","Supplier Name","Recipient GSTIN/UIN","Recipient Name","Original Document Number","Original Document Date","Revised Document Number","Revised Document Date","Value of supplies made","Place Of Supply","Document type","Rate","Taxable Value","Cess Amount"],"w":[44.5,21.2,27.5,20.7,27.5,23.5,27.5,23.5,22.8,24.3,20.2,14.7,21.5,23.0]},{"n":"ecoab2c","r1":["Summary For Supplies U/s 9(5)-15A-B2C","Original details ",null,null,"Revised details ",null,null,"HELP"],"r2":[null,null,"No. of Supplier",null,null,null,"Total Taxable Value","Total Cess"],"h":["Financial Year","Original Month","Supplier GSTIN/UIN","Supplier Name","Place Of Supply","Rate","Taxable Value","Cess Amount"],"w":[43.3,15.3,29.2,21.3,31.7,12.8,26.5,26.0]},{"n":"ecoaurp2b","r1":["Summary For Supplies U/s 9(5)-15A-URP2B","Original details ",null,null,"Revised details ",null,null,null,null,null,null,"HELP"],"r2":["No. of Recipients",null,"No. of Documents",null,null,null,"Total value of supplies made",null,null,null,"Total Taxable Value","Total Cess"],"h":["Recipient GSTIN/UIN","Recipient Name","Original Document Number","Original Document Date","Revised Document Number","Revised Document Date","Value of supplies made","Document type","Place Of Supply","Rate","Taxable Value","Cess Amount"],"w":[46.3,15.8,27.5,23.5,27.5,23.5,30.3,20.7,23.7,10.2,23.3,27.2]},{"n":"ecoaurp2c","r1":["Summary For Supplies U/s 9(5)-15A-URP2C","Original details ","Revised details ",null,null,"HELP"],"r2":[null,null,null,null,"Total Taxable Value","Total Cess"],"h":["Financial Year","Original Month","Place Of Supply","Rate","Taxable Value","Cess Amount"],"w":[46.5,17.5,18.5,10.5,22.7,23.2]}],"uqc":["BAG-BAGS","BAL-BALE","BDL-BUNDLES","BKL-BUCKLES","BOU-BILLION OF UNITS","BOX-BOX","BTL-BOTTLES","BUN-BUNCHES","CAN-CANS","CBM-CUBIC METERS","CCM-CUBIC CENTIMETERS","CMS-CENTIMETERS","CTN-CARTONS","DOZ-DOZENS","DRM-DRUMS","GGK-GREAT GROSS","GMS-GRAMMES","GRS-GROSS","GYD-GROSS YARDS","KGS-KILOGRAMS","KLR-KILOLITRE","KME-KILOMETRE","LTR-LITRES","MLT-MILILITRE","MTR-METERS","MTS-METRIC TON","NOS-NUMBERS","PAC-PACKS","PCS-PIECES","PRS-PAIRS","QTL-QUINTAL","ROL-ROLLS","SET-SETS","SQF-SQUARE FEET","SQM-SQUARE METERS","SQY-SQUARE YARDS","TBS-TABLETS","TGM-TEN GROSS","THD-THOUSANDS","TON-TONNES","TUB-TUBES","UGS-US GALLONS","UNT-UNITS","YDS-YARDS","OTH-OTHERS"]};

  var STATES = [
    ['01', 'Jammu & Kashmir', 'JK'], ['02', 'Himachal Pradesh', 'HP'], ['03', 'Punjab', 'PB'],
    ['04', 'Chandigarh', 'CH'], ['05', 'Uttarakhand', 'UK'], ['06', 'Haryana', 'HR'], ['07', 'Delhi', 'DL'],
    ['08', 'Rajasthan', 'RJ'], ['09', 'Uttar Pradesh', 'UP'], ['10', 'Bihar', 'BR'], ['11', 'Sikkim', 'SK'],
    ['12', 'Arunachal Pradesh', 'AR'], ['13', 'Nagaland', 'NL'], ['14', 'Manipur', 'MN'], ['15', 'Mizoram', 'MZ'],
    ['16', 'Tripura', 'TR'], ['17', 'Meghalaya', 'ML'], ['18', 'Assam', 'AS'], ['19', 'West Bengal', 'WB'],
    ['20', 'Jharkhand', 'JH'], ['21', 'Odisha', 'OD'], ['22', 'Chhattisgarh', 'CG'], ['23', 'Madhya Pradesh', 'MP'],
    ['24', 'Gujarat', 'GJ'], ['25', 'Daman & Diu', 'DD'], ['26', 'Dadra & Nagar Haveli & Daman & Diu', 'DN'],
    ['27', 'Maharashtra', 'MH'], ['29', 'Karnataka', 'KA'], ['30', 'Goa', 'GA'], ['31', 'Lakshdweep', 'LD'],
    ['32', 'Kerala', 'KL'], ['33', 'Tamil Nadu', 'TN'], ['34', 'Puducherry', 'PY'],
    ['35', 'Andaman & Nicobar Islands', 'AN'], ['36', 'Telangana', 'TS'], ['37', 'Andhra Pradesh', 'AP'],
    ['38', 'Ladakh', 'LA'], ['96', 'Foreign Country', 'FC'], ['97', 'Other Territory', 'OT']
  ];
  var STATE_BY_CODE = {}, STATE_LOOKUP = {};
  STATES.forEach(function (s) {
    STATE_BY_CODE[s[0]] = s;
    STATE_LOOKUP[s[0]] = s[0]; STATE_LOOKUP[String(+s[0])] = s[0];
    STATE_LOOKUP[norm(s[1])] = s[0]; STATE_LOOKUP[norm(s[2])] = s[0];
  });
  STATE_LOOKUP.orissa = '21'; STATE_LOOKUP.pondicherry = '34'; STATE_LOOKUP.lakshadweep = '31';
  STATE_LOOKUP.tg = '36'; STATE_LOOKUP.or = '21'; STATE_LOOKUP.ut = '05'; STATE_LOOKUP.uttaranchal = '05';

  var RATES = [0, 0.1, 0.25, 1, 1.5, 3, 5, 6, 7.5, 12, 18, 28, 40];

  var UQC_LIST = TEMPLATE.uqc; // "PCS-PIECES" ...
  var UQC_BY_CODE = {};
  UQC_LIST.forEach(function (u) { UQC_BY_CODE[u.split('-')[0]] = u; });
  var UQC_ALIAS = { PC: 'PCS', PIECE: 'PCS', PIECES: 'PCS', NO: 'NOS', NOS: 'NOS', NUMBER: 'NOS', NUMBERS: 'NOS',
    KG: 'KGS', KGS: 'KGS', KILOGRAM: 'KGS', KILOGRAMS: 'KGS', LTR: 'LTR', LITRE: 'LTR', LITRES: 'LTR', LITER: 'LTR',
    MTR: 'MTR', METER: 'MTR', METERS: 'MTR', METRE: 'MTR', SET: 'SET', SETS: 'SET', BOX: 'BOX', BOXES: 'BOX',
    PAC: 'PAC', PACK: 'PAC', PACKS: 'PAC', PKT: 'PAC', DOZ: 'DOZ', DOZEN: 'DOZ', PAIR: 'PRS', PAIRS: 'PRS', PRS: 'PRS',
    UNIT: 'UNT', UNITS: 'UNT', GM: 'GMS', GMS: 'GMS', GRAM: 'GMS', GRAMS: 'GMS', TON: 'TON', TONNE: 'TON',
    BAG: 'BAG', BAGS: 'BAG', BTL: 'BTL', BOTTLE: 'BTL', BOTTLES: 'BTL', CTN: 'CTN', CARTON: 'CTN', ROL: 'ROL', ROLL: 'ROL' };

  var DOC_NATURE = {
    INV: { num: 1, typ: 'Invoices for outward supply' },
    DN: { num: 4, typ: 'Debit Note' },
    CN: { num: 5, typ: 'Credit Note' },
    RV: { num: 6, typ: 'Receipt Voucher' }
  };

  var EXEMP_ROWS = [
    ['INTRB2B', 'Inter-State supplies to registered persons'],
    ['INTRAB2B', 'Intra-State supplies to registered persons'],
    ['INTRB2C', 'Inter-State supplies to unregistered persons'],
    ['INTRAB2C', 'Intra-State supplies to unregistered persons']
  ];

  var INV_TYPES = {
    R: { label: 'Regular B2B', igst: false, nopay: false },
    SEWP: { label: 'SEZ supplies with payment', igst: true, nopay: false },
    SEWOP: { label: 'SEZ supplies without payment', igst: true, nopay: true },
    DE: { label: 'Deemed Exp', igst: false, nopay: false },
    IGST7: { label: 'Intra-State supplies attracting IGST', igst: true, nopay: false, json: 'R' }
  };

  /* ------------------------------------------------------------------ *
   * 2. Small helpers
   * ------------------------------------------------------------------ */

  function norm(s) { return String(s == null ? '' : s).toLowerCase().replace(/[^a-z0-9]/g, ''); }
  function str(v) { return v == null ? '' : String(v).trim(); }
  function r2(x) { var n = Math.round((Math.abs(x) + 1e-9) * 100) / 100; return x < 0 ? -n : n; }
  function num(v) {
    if (v == null || v === '') return 0;
    if (typeof v === 'number') return isFinite(v) ? v : 0;
    var s = String(v).replace(/[,₹\s]/g, '');
    if (/^\(.*\)$/.test(s)) s = '-' + s.slice(1, -1);
    var n = parseFloat(s);
    return isFinite(n) ? n : 0;
  }
  function isNumLike(v) {
    if (typeof v === 'number') return isFinite(v);
    return /^\s*-?[\d,]*\.?\d+\s*$/.test(String(v == null ? '' : v));
  }
  function pad2(n) { return (n < 10 ? '0' : '') + n; }
  var MON = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'];
  var MON_IDX = {}; MON.forEach(function (m, i) { MON_IDX[m.toLowerCase()] = i + 1; });

  // Date represented as {y,m,d}; helpers avoid any time-zone shifts.
  function serialToDate(n, date1904) {
    var days = Math.floor(n) + (date1904 ? 1462 : 0);
    if (days < 61) days += 1; // Excel's phantom 29-Feb-1900
    var t = Date.UTC(1899, 11, 30) + days * 86400000;
    var d = new Date(t);
    return { y: d.getUTCFullYear(), m: d.getUTCMonth() + 1, d: d.getUTCDate() };
  }
  function parseDate(v, date1904) {
    if (v == null || v === '') return null;
    if (typeof v === 'number') return v > 0 ? serialToDate(v, date1904) : null;
    var s = String(v).trim(), m;
    if ((m = /^(\d{4})-(\d{1,2})-(\d{1,2})/.exec(s))) return { y: +m[1], m: +m[2], d: +m[3] };
    if ((m = /^(\d{1,2})[-\/.](\d{1,2})[-\/.](\d{2,4})$/.exec(s))) return { y: yr(+m[3]), m: +m[2], d: +m[1] };
    if ((m = /^(\d{1,2})[-\/ ]([A-Za-z]{3})[A-Za-z]*[-\/ ,]+(\d{2,4})$/.exec(s)) && MON_IDX[m[2].toLowerCase()])
      return { y: yr(+m[3]), m: MON_IDX[m[2].toLowerCase()], d: +m[1] };
    if (/^\d+(\.\d+)?$/.test(s)) return serialToDate(+s, date1904);
    return null;
  }
  function yr(y) { return y < 100 ? 2000 + y : y; }
  function dmy(d) { return d ? pad2(d.d) + '-' + pad2(d.m) + '-' + d.y : ''; }
  function dMonY(d) { return d ? pad2(d.d) + '-' + MON[d.m - 1] + '-' + d.y : ''; }
  function dateSerial(d) { return (Date.UTC(d.y, d.m - 1, d.d) - Date.UTC(1899, 11, 30)) / 86400000; }
  function dateKey(d) { return d ? d.y * 10000 + d.m * 100 + d.d : 0; }

  function stateCode(v) {
    var s = str(v);
    if (!s) return '';
    var m = /^(\d{1,2})\b/.exec(s);
    if (m) { var c = pad2(+m[1]); return STATE_BY_CODE[c] ? c : ''; }
    var k = norm(s.replace(/^\d+\s*-\s*/, ''));
    return STATE_LOOKUP[k] || '';
  }
  function posLabel(code) { var s = STATE_BY_CODE[code]; return s ? s[0] + '-' + s[1] : code; }

  var GSTIN_CHARS = '0123456789ABCDEFGHIJKLMNOPQRSTUVWXYZ';
  function gstinChecksumOk(g) {
    if (!/^[0-9A-Z]{15}$/.test(g)) return false;
    var sum = 0;
    for (var i = 0; i < 14; i++) {
      var v = GSTIN_CHARS.indexOf(g[i]) * (i % 2 ? 2 : 1);
      sum += Math.floor(v / 36) + (v % 36);
    }
    return GSTIN_CHARS[(36 - (sum % 36)) % 36] === g[14];
  }
  function gstinShapeOk(g) {
    return /^\d{2}[A-Z]{5}\d{4}[A-Z][0-9A-Z]Z[0-9A-Z]$/.test(g) || /^\d{4}[A-Z]{3}\d{5}[0-9A-Z]{3}$/.test(g)
      || /^\d{2}[A-Z0-9]{13}$/.test(g);
  }

  function uqcFull(v) {
    var s = str(v).toUpperCase();
    if (!s || s === 'NA' || s === 'N/A') return '';
    var code = s.split(/[-\s]/)[0];
    if (UQC_BY_CODE[code]) return UQC_BY_CODE[code];
    var alias = UQC_ALIAS[s.replace(/[^A-Z]/g, '')] || UQC_ALIAS[code];
    if (alias) return UQC_BY_CODE[alias];
    for (var i = 0; i < UQC_LIST.length; i++) {
      if (norm(UQC_LIST[i].split('-').slice(1).join('-')) === norm(s)) return UQC_LIST[i];
    }
    return null; // unknown
  }

  function isService(hsn) { return /^99/.test(hsn); }

  /* ------------------------------------------------------------------ *
   * 3. XLSX reader (zip + streaming-ish XML scan; avoids heavy libs)
   * ------------------------------------------------------------------ */

  var ENT = { amp: '&', lt: '<', gt: '>', quot: '"', apos: "'" };
  function unesc(s) {
    if (s.indexOf('&') >= 0) {
      s = s.replace(/&(#x[0-9a-fA-F]+|#\d+|amp|lt|gt|quot|apos);/g, function (_, e) {
        if (e[0] === '#') return String.fromCodePoint(e[1] === 'x' ? parseInt(e.slice(2), 16) : +e.slice(1));
        return ENT[e];
      });
    }
    if (s.indexOf('_x') >= 0) s = s.replace(/_x([0-9A-Fa-f]{4})_/g, function (_, h) { return String.fromCharCode(parseInt(h, 16)); });
    return s;
  }
  function attr(tag, name) {
    var m = new RegExp('\\s' + name + '="([^"]*)"').exec(tag);
    return m ? unesc(m[1]) : null;
  }
  function decode(u8) {
    if (typeof TextDecoder !== 'undefined') return new TextDecoder('utf-8').decode(u8);
    return fflate.strFromU8(u8);
  }
  function textOfRich(x) {
    // concatenate <t> runs, skipping phonetic runs
    x = x.replace(/<rPh\b[\s\S]*?<\/rPh>/g, '');
    var out = '', re = /<t(?:\s[^>]*)?>([\s\S]*?)<\/t>|<t\s*\/>/g, m;
    while ((m = re.exec(x))) out += m[1] || '';
    return unesc(out);
  }
  function colIndex(ref) {
    var n = 0;
    for (var i = 0; i < ref.length; i++) {
      var c = ref.charCodeAt(i);
      if (c < 65 || c > 90) break;
      n = n * 26 + (c - 64);
    }
    return n - 1;
  }

  function openXlsx(u8) {
    var files = fflate.unzipSync(u8, { filter: function (f) {
      return /^xl\/(workbook\.xml|_rels\/workbook\.xml\.rels|sharedStrings\.xml)$/i.test(f.name) || /^\[Content_Types\]/.test(f.name);
    } });
    var wbXml = files['xl/workbook.xml'] ? decode(files['xl/workbook.xml']) : null;
    if (!wbXml) throw new Error('This does not look like an .xlsx workbook (xl/workbook.xml missing).');
    var relsXml = decode(files['xl/_rels/workbook.xml.rels'] || new Uint8Array());
    var rels = {}, m, re = /<Relationship\b[^>]*>/g;
    while ((m = re.exec(relsXml))) rels[attr(m[0], 'Id')] = attr(m[0], 'Target');
    var sheets = [];
    re = /<sheet\b[^>]*>/g;
    while ((m = re.exec(wbXml))) {
      var target = rels[attr(m[0], 'r:id')] || '';
      target = target.replace(/^\//, '');
      if (!/^xl\//.test(target)) target = 'xl/' + target;
      sheets.push({ name: attr(m[0], 'name'), path: target });
    }
    var date1904 = /date1904="(1|true)"/.test(wbXml);
    var sst = [];
    if (files['xl/sharedStrings.xml']) {
      var sx = decode(files['xl/sharedStrings.xml']);
      var si = /<si>([\s\S]*?)<\/si>|<si\/>/g;
      while ((m = si.exec(sx))) sst.push(m[1] ? textOfRich(m[1]) : '');
    }
    return {
      sheets: sheets,
      date1904: date1904,
      readSheet: function (name, maxRows) {
        var sh = sheets.filter(function (s) { return s.name === name; })[0];
        if (!sh) throw new Error('Sheet not found: ' + name);
        var f = fflate.unzipSync(u8, { filter: function (x) { return x.name === sh.path; } });
        if (!f[sh.path]) throw new Error('Sheet part missing: ' + sh.path);
        return parseSheetXml(decode(f[sh.path]), sst, maxRows);
      }
    };
  }

  function parseSheetXml(xml, sst, maxRows) {
    var rows = [];
    var start = xml.indexOf('<sheetData');
    if (start < 0) return rows;
    var end = xml.lastIndexOf('</sheetData>');
    var body = xml.slice(start, end < 0 ? xml.length : end);
    var rowRe = /<row\b([^>]*?)(\/>|>([\s\S]*?)<\/row>)/g, rm, nextRow = 0;
    var cellRe = /<c\b([^>]*?)(\/>|>([\s\S]*?)<\/c>)/g;
    while ((rm = rowRe.exec(body))) {
      var rAttr = attr(rm[1], 'r');
      var ri = rAttr ? (+rAttr - 1) : nextRow;
      nextRow = ri + 1;
      if (maxRows && ri >= maxRows) break;
      var arr = [];
      var inner = rm[3];
      if (inner) {
        cellRe.lastIndex = 0;
        var cm, nextCol = 0;
        while ((cm = cellRe.exec(inner))) {
          var a = cm[1];
          var ref = attr(a, 'r');
          var ci = ref ? colIndex(ref) : nextCol;
          nextCol = ci + 1;
          var c = cm[3];
          if (!c) continue;
          var t = attr(a, 't');
          var val = null, vm;
          if (t === 'inlineStr') {
            val = textOfRich(c);
          } else if ((vm = /<v>([\s\S]*?)<\/v>/.exec(c))) {
            var raw = vm[1];
            if (t === 's') val = sst[+raw];
            else if (t === 'str') val = unesc(raw);
            else if (t === 'b') val = raw === '1';
            else if (t === 'e') val = null;
            else if (t === 'd') val = raw;
            else { val = +raw; if (!isFinite(val)) val = unesc(raw); }
          }
          if (val !== null && val !== undefined && val !== '') arr[ci] = val;
        }
      }
      rows[ri] = arr;
    }
    return rows;
  }

  // Minimal RFC-4180 CSV parser for CSV exports.
  function parseCsv(text) {
    if (text.charCodeAt(0) === 0xFEFF) text = text.slice(1);
    var rows = [], row = [], field = '', i = 0, q = false, n = text.length;
    var sep = (text.split('\n')[0].split('\t').length > text.split('\n')[0].split(',').length) ? '\t' : ',';
    while (i < n) {
      var ch = text[i];
      if (q) {
        if (ch === '"') { if (text[i + 1] === '"') { field += '"'; i += 2; continue; } q = false; i++; continue; }
        field += ch; i++; continue;
      }
      if (ch === '"') { q = true; i++; continue; }
      if (ch === sep) { row.push(field); field = ''; i++; continue; }
      if (ch === '\r') { i++; continue; }
      if (ch === '\n') { row.push(field); rows.push(row); row = []; field = ''; i++; continue; }
      field += ch; i++;
    }
    if (field !== '' || row.length) { row.push(field); rows.push(row); }
    return rows.map(function (r) {
      return r.map(function (v) { v = v.trim(); return v === '' ? undefined : (isNumLike(v) && !/^0\d/.test(v) ? +v.replace(/,/g, '') : v); });
    });
  }

  /* ------------------------------------------------------------------ *
   * 4. Input mapping (Munim "GSTR-1 Detail" layout, matched by header text)
   * ------------------------------------------------------------------ */

  var FIELDS = [
    ['docType', function (h) { return h === 'documenttypecode' || h === 'documenttype'; }, true],
    ['invDate', function (h) { return h === 'invoicedate'; }, true],
    ['invNo', function (h) { return h === 'invoiceno' || h === 'invoicenumber'; }, true],
    ['name', function (h) { return h === 'billingname' || h === 'customername' || h === 'receivername'; }],
    ['gstin', function (h) { return h === 'billinggstin' || h === 'gstinuin' || h === 'customergstin' || h === 'gstinuinofrecipient'; }],
    ['noteDate', function (h) { return h === 'crdbnotedate' || h === 'notedate'; }],
    ['noteNo', function (h) { return h === 'crdbnoteno' || h === 'noteno' || h === 'notenumber'; }],
    ['noteType', function (h) { return /notetype$/.test(h); }],
    ['pos', function (h) { return h.indexOf('placeofsupply') === 0; }, true],
    ['desc', function (h) { return h === 'itemdescription' || h === 'description'; }],
    ['hsn', function (h) { return h.indexOf('hsn') === 0; }, true],
    ['qty', function (h) { return h === 'quantity' || h === 'qty'; }],
    ['uqc', function (h) { return h.indexOf('itemunit') === 0 || h === 'uom' || h === 'uqc' || h === 'unit'; }],
    ['txval', function (h) { return h === 'taxablevalue'; }, true],
    ['rate', function (h) { return h === 'gstrate' || h === 'rate' || h === 'taxrate'; }, true],
    ['cess', function (h) { return h === 'cessamount' || h === 'cess'; }],
    ['rchrg', function (h) { return h.indexOf('reversecharge') >= 0; }],
    ['gstType', function (h) { return h === 'itemgsttype' || h.indexOf('nilratedexempt') >= 0; }],
    ['invType', function (h) { return h === 'invoicetype'; }],
    ['urType', function (h) { return h === 'urtype'; }],
    ['expType', function (h) { return h === 'typeofexport' || h === 'exporttype'; }],
    ['portCode', function (h) { return h.indexOf('shippingportcode') === 0 || h === 'portcode'; }],
    ['sbNum', function (h) { return h.indexOf('shippingbillno') === 0 || h.indexOf('shippingbillnumber') === 0; }],
    ['sbDate', function (h) { return h.indexOf('shippingbilldate') === 0; }],
    ['origInvDate', function (h) { return h.indexOf('originalinvoicedate') === 0; }],
    ['origInvNo', function (h) { return h.indexOf('originalinvoicenumber') === 0 || h.indexOf('originalinvoiceno') === 0; }],
    ['origNoteDate', function (h) { return h.indexOf('originalcreditdebitnotedate') === 0; }],
    ['origNoteNo', function (h) { return h.indexOf('originalcreditdebitnotenumber') === 0 || h.indexOf('originalcreditdebitnoteno') === 0; }],
    ['advAdj', function (h) { return h.indexOf('adjustmentamount') === 0; }],
    ['diffPct', function (h) { return h === 'applicableoftaxrate' || h === 'applicabletaxrate'; }],
    ['val', function (h) { return h.indexOf('totaltransactionvalue') === 0 || h === 'invoicevalue'; }],
    ['cancelled', function (h) { return h.indexOf('isinvoicecancelled') === 0 || h === 'cancelled'; }],
    ['etin', function (h) { return h === 'ecommercegstin' || h === 'etin'; }]
  ];

  function findHeader(rows) {
    for (var i = 0; i < Math.min(rows.length, 15); i++) {
      var r = rows[i] || [];
      var hs = [];
      for (var c = 0; c < r.length; c++) hs[c] = norm(r[c]);
      if (hs.indexOf('documenttypecode') >= 0 || (hs.indexOf('taxablevalue') >= 0 && hs.some(function (h) { return h && h.indexOf('hsn') === 0; }))) {
        var map = {}, missing = [];
        FIELDS.forEach(function (f) {
          for (var c = 0; c < hs.length; c++) if (hs[c] && f[1](hs[c])) { if (!(f[0] in map)) map[f[0]] = c; }
          if (f[2] && !(f[0] in map)) missing.push(f[0]);
        });
        return { row: i, map: map, missing: missing, headers: r };
      }
    }
    return null;
  }

  function pickSheet(book) {
    var best = null;
    for (var i = 0; i < book.sheets.length; i++) {
      var s = book.sheets[i];
      var head = book.readSheet(s.name, 15);
      var h = findHeader(head);
      if (h && h.missing.length === 0) return s.name;
      if (h && !best) best = s.name;
    }
    return best;
  }

  /* Read workbook/CSV bytes into normalised line objects. */
  function loadSource(u8, fileName) {
    var rows, sheetName = null, date1904 = false;
    if (/\.(csv|txt|tsv)$/i.test(fileName || '')) {
      rows = parseCsv(decode(u8));
      sheetName = '(CSV)';
    } else {
      if (!(u8[0] === 0x50 && u8[1] === 0x4B)) throw new Error('Unsupported file. Please upload the Munim export as .xlsx (or .csv). Old .xls files must be re-saved as .xlsx.');
      var book = openXlsx(u8);
      date1904 = book.date1904;
      sheetName = pickSheet(book);
      if (!sheetName) throw new Error('Could not find the "GSTR-1 Detail" sheet (no header row with "Document Type Code" / "Taxable Value" / "HSN").');
      rows = book.readSheet(sheetName);
    }
    var hdr = findHeader(rows);
    if (!hdr) throw new Error('Header row not found.');
    if (hdr.missing.length) throw new Error('Required column(s) missing: ' + hdr.missing.join(', '));
    var M = hdr.map, lines = [];
    var DATE_F = { invDate: 1, noteDate: 1, sbDate: 1, origInvDate: 1, origNoteDate: 1 };
    for (var i = hdr.row + 1; i < rows.length; i++) {
      var r = rows[i];
      if (!r || !r.length) continue;
      var o = { _row: i + 1 };
      var any = false;
      for (var k in M) {
        var v = r[M[k]];
        if (v === undefined || v === null || v === '') { o[k] = null; continue; }
        if (DATE_F[k]) o[k] = parseDate(v, date1904);
        else if (typeof v === 'string' && v.charAt(0) === '=') o[k] = null; // uncached formula
        else o[k] = v;
        any = true;
      }
      if (!any || !str(o.docType) && !str(o.invNo) && !str(o.noteNo)) continue;
      lines.push(o);
    }
    return { sheet: sheetName, headerRow: hdr.row + 1, columns: Object.keys(M), lines: lines };
  }

  /* ------------------------------------------------------------------ *
   * 5. Segregation engine
   * ------------------------------------------------------------------ */

  function defaultOptions() {
    return {
      supplierGstin: '',
      homeState: '27',
      period: '',              // MMYYYY
      b2clLimit: 100000,       // invoice value above which inter-state B2C invoices go to B2CL
      zeroRateAsNil: false,    // 0% lines without an explicit Item GST Type -> Table 8 (nil rated)
      reduceQtyOnCN: true,     // credit notes reduce HSN quantity
      mergeNoteUnits: true,    // one row per HSN + rate: fold rows booked in other UQCs into the main unit
      gapsAsCancelled: false,  // Table 13: numbers missing inside a series are treated as cancelled
      hsnDesc: false,          // fill HSN description from item descriptions (otherwise left for the GSTN HSN master)
      jsonVersion: 'GST3.2.3',
      unregGstins: [],         // recipient GSTINs to report as unregistered (e.g. cancelled registrations)
      hsnMap: {},              // { wrongHsn: correctHsn } corrections applied before aggregation
      invalidHsn: []           // HSN codes the portal rejected (from an error report) — flagged until corrected
    };
  }

  function docKind(s) {
    var n = norm(s);
    if (n === 'invoice' || n === 'inv' || n === 'taxinvoice' || n === 'billofsupply') return 'INV';
    if (n === 'creditnote' || n === 'cn' || n === 'c') return 'CN';
    if (n === 'debitnote' || n === 'dn' || n === 'd') return 'DN';
    if (n === 'advanceadjustment' || n === 'advanceadjusted') return 'ADJ';
    if (n === 'advance' || n === 'advancereceipt' || n === 'advancereceived') return 'ADV';
    return '';
  }
  function invTypeCode(s) {
    var n = norm(s);
    if (!n || n === 'regularb2b' || n === 'regular' || n === 'r' || n === 'b2b') return 'R';
    if (n === 'sezsupplieswithpayment' || n === 'sewp' || n === 'sezwithpayment') return 'SEWP';
    if (n === 'sezsupplieswithoutpayment' || n === 'sewop' || n === 'sezwithoutpayment') return 'SEWOP';
    if (n === 'deemedexp' || n === 'deemedexport' || n === 'de') return 'DE';
    if (n === 'intrastatesuppliesattractingigst') return 'IGST7';
    return null;
  }
  function gstTypeCode(s) {
    var n = norm(s);
    if (!n) return '';
    if (n.indexOf('nil') === 0) return 'NIL';
    if (n.indexOf('exempt') === 0) return 'EXEMPT';
    if (n.indexOf('nongst') === 0 || n === 'nongstsupply') return 'NONGST';
    return '';
  }
  function expTypeCode(s) {
    var n = norm(s);
    if (n === 'expwp' || n === 'wpay' || n === 'withpayment') return 'WPAY';
    if (n === 'expwop' || n === 'wopay' || n === 'withoutpayment') return 'WOPAY';
    return '';
  }
  function yes(v) { var n = norm(v); return n === 'y' || n === 'yes' || n === 'true' || n === '1'; }

  function Issues() { this.map = {}; this.list = []; }
  Issues.prototype.add = function (level, code, msg, ref) {
    var k = code;
    var it = this.map[k];
    if (!it) { it = this.map[k] = { level: level, code: code, msg: msg, count: 0, refs: [] }; this.list.push(it); }
    it.count++;
    if (ref && it.refs.length < 25 && it.refs.indexOf(ref) < 0) it.refs.push(ref);
  };
  Issues.prototype.sorted = function () {
    var o = { error: 0, warn: 1, info: 2 };
    return this.list.slice().sort(function (a, b) { return o[a.level] - o[b.level] || b.count - a.count; });
  };

  function newAmt() { return { txval: 0, iamt: 0, camt: 0, samt: 0, csamt: 0 }; }
  function addAmt(a, b, sign) { sign = sign || 1; a.txval += sign * b.txval; a.iamt += sign * b.iamt; a.camt += sign * b.camt; a.samt += sign * b.samt; a.csamt += sign * b.csamt; }

  function build(src, opt) {
    opt = Object.assign(defaultOptions(), opt || {});
    var issues = new Issues();
    var home = opt.homeState;
    var supplierGstin = str(opt.supplierGstin).toUpperCase();
    if (supplierGstin && /^\d{2}/.test(supplierGstin)) home = supplierGstin.slice(0, 2);
    var unreg = {}, hsnMap = {}, invalidHsn = {};
    (opt.unregGstins || []).forEach(function (g) { g = str(g).toUpperCase(); if (g) unreg[g] = 1; });
    Object.keys(opt.hsnMap || {}).forEach(function (k) { var v = str(opt.hsnMap[k]).replace(/\s/g, ''); if (/^\d{4,8}$/.test(v)) hsnMap[str(k)] = v; });
    (opt.invalidHsn || []).forEach(function (h) { invalidHsn[str(h)] = 1; });

    // ---- group lines into documents
    var docs = {}, order = [];
    var amendRows = 0;
    src.lines.forEach(function (L) {
      var kind = docKind(L.docType);
      if (!kind) { issues.add('error', 'doctype', 'Unrecognised Document Type Code — row skipped', 'Row ' + L._row + ': ' + str(L.docType)); return; }
      var isNote = kind === 'CN' || kind === 'DN';
      if (isNote && L.noteType) { var nt = norm(L.noteType); if (nt === 'd') kind = 'DN'; else if (nt === 'c') kind = 'CN'; }
      var no = str(isNote ? L.noteNo : L.invNo);
      var dt = isNote ? L.noteDate : L.invDate;
      if (!no && (kind === 'ADV' || kind === 'ADJ')) no = '(adv-row-' + L._row + ')';
      if (!no) { issues.add('error', 'nodocno', 'Document number missing — row skipped', 'Row ' + L._row); return; }
      if (L.origInvNo || L.origNoteNo) amendRows++;
      var key = kind + '|' + no;
      var d = docs[key];
      if (!d) {
        d = docs[key] = { kind: kind, no: no, date: dt, gstin: str(L.gstin).toUpperCase().replace(/\s/g, ''), name: str(L.name),
          posRaw: str(L.pos), pos: stateCode(L.pos), lines: [], valGiven: null, cancelled: false,
          rchrg: yes(L.rchrg) ? 'Y' : 'N', invTypeRaw: str(L.invType), urType: str(L.urType).toUpperCase(),
          expType: expTypeCode(L.expType), portCode: str(L.portCode), sbNum: str(L.sbNum), sbDate: L.sbDate,
          etin: str(L.etin).toUpperCase(), diffPct: num(L.diffPct), firstRow: L._row };
        order.push(key);
      } else {
        if ((d.gstin !== str(L.gstin).toUpperCase().replace(/\s/g, '')) || d.pos !== stateCode(L.pos) || dateKey(d.date) !== dateKey(dt))
          issues.add('error', 'dochdr', 'Same document number has different GSTIN / place of supply / date on different lines (possible duplicate number)', kind + ' ' + no);
      }
      if (yes(L.cancelled)) d.cancelled = true;
      if (L.val != null && isNumLike(L.val)) d.valGiven = num(L.val);
      d.lines.push(L);
    });
    if (amendRows) issues.add('warn', 'amend', 'Amendment columns are filled on ' + amendRows + ' line(s). Amendment tables (B2BA, CDNRA, etc.) are not generated — enter those directly in the offline tool.', '');

    // ---- period detection
    var monthCount = {};
    order.forEach(function (k) { var d = docs[k].date; if (d) { var p = pad2(d.m) + d.y; monthCount[p] = (monthCount[p] || 0) + 1; } });
    var detectedPeriod = Object.keys(monthCount).sort(function (a, b) { return monthCount[b] - monthCount[a]; })[0] || '';
    var period = opt.period || detectedPeriod;

    // ---- outputs
    var out = {
      b2b: [], b2cl: [], exp: [], cdnr: [], cdnur: [],
      b2cs: {}, at: {}, atadj: {},
      nil: { INTRB2B: [0, 0, 0], INTRAB2B: [0, 0, 0], INTRB2C: [0, 0, 0], INTRAB2C: [0, 0, 0] },
      hsn: { b2b: {}, b2c: {} },
      series: {}
    };
    var src_tot = { INV: newAmt(), CN: newAmt(), DN: newAmt(), ADV: newAmt(), ADJ: newAmt() };
    var stats = { lines: src.lines.length, docs: order.length, byKind: { INV: 0, CN: 0, DN: 0, ADV: 0, ADJ: 0 }, cancelled: 0, zeroLines: 0 };
    var invNumbers = {}, noteNumbers = {};

    order.forEach(function (key) {
      var d = docs[key];
      stats.byKind[d.kind]++;
      var ref = (d.kind === 'INV' ? 'Inv ' : d.kind === 'CN' ? 'CN ' : d.kind === 'DN' ? 'DN ' : 'Adv ') + d.no;

      // Table 13 series registration (before cancellation check)
      if (d.kind === 'INV' || d.kind === 'CN' || d.kind === 'DN' || (d.kind === 'ADV' && d.no.charAt(0) !== '(')) {
        registerSeries(out.series, d.kind === 'ADV' ? 'RV' : d.kind, d.no, d.cancelled);
      }
      if (d.kind === 'INV') invNumbers[d.no] = 1; else if (d.kind === 'CN' || d.kind === 'DN') noteNumbers[d.no] = 1;
      if (d.cancelled) { stats.cancelled++; return; }

      // --- document-level validation
      if (d.kind !== 'ADV' && d.kind !== 'ADJ') {
        if (d.no.length > 16) issues.add('error', 'docnolen', 'Document number longer than 16 characters', ref);
        if (!/^[A-Za-z0-9\/-]+$/.test(d.no)) issues.add('error', 'docnochars', 'Document number has characters other than letters, digits, "/" and "-"', ref);
        if (/^0+$/.test(d.no.replace(/[\/-]/g, ''))) issues.add('error', 'docnozero', 'Document number cannot be all zeros', ref);
        if (!d.date) issues.add('error', 'nodate', 'Document date missing or unreadable', ref);
        else if (period && pad2(d.date.m) + d.date.y !== period) {
          var pk = +period.slice(2) * 100 + +period.slice(0, 2), dk = d.date.y * 100 + d.date.m;
          if (dk > pk) issues.add('error', 'futuredate', 'Document dated after the return period', ref + ' (' + dMonY(d.date) + ')');
          else if (d.kind === 'INV') issues.add('warn', 'olddate', 'Invoice dated before the return period (allowed, but check it was not reported earlier)', ref + ' (' + dMonY(d.date) + ')');
        }
      }
      if (!d.pos) {
        if (d.posRaw) issues.add('error', 'posbad', 'Place of supply not recognised', ref + ': ' + d.posRaw);
        else issues.add('error', 'posmissing', 'Place of supply missing', ref);
      }
      if (d.gstin && unreg[d.gstin]) {
        issues.add('info', 'unreg', 'Recipient GSTIN marked as cancelled / unregistered under Corrections — reported as a B2C supply', ref + ': ' + d.gstin);
        d.gstin = '';
      }
      var registered = !!d.gstin;
      if (registered) {
        if (!gstinShapeOk(d.gstin) || (/^\d{2}[A-Z]{5}\d{4}/.test(d.gstin) && !gstinChecksumOk(d.gstin)))
          issues.add('error', 'gstin', 'Recipient GSTIN/UIN is invalid (format or check digit)', ref + ': ' + d.gstin);
        if (supplierGstin && d.gstin === supplierGstin) issues.add('warn', 'selfgstin', 'Recipient GSTIN is the same as the supplier GSTIN', ref);
      }
      if (d.etin && !gstinShapeOk(d.etin)) issues.add('error', 'etin', 'E-commerce GSTIN looks invalid', ref + ': ' + d.etin);

      var itCode = invTypeCode(d.invTypeRaw);
      if (itCode === null) { issues.add('error', 'invtype', 'Invoice Type not recognised — treated as Regular B2B', ref + ': ' + d.invTypeRaw); itCode = 'R'; }
      var it = INV_TYPES[itCode];

      // --- section decision
      var sect;
      if (d.kind === 'ADV') sect = 'AT';
      else if (d.kind === 'ADJ') sect = 'ATADJ';
      else if (d.kind === 'INV') {
        if (d.expType || (!registered && d.pos === '96')) {
          sect = 'EXP';
          if (!d.expType) { d.expType = 'WOPAY'; issues.add('warn', 'exptype', 'Place of supply is Foreign Country but Type of Export is blank — treated as export without payment (WOPAY)', ref); }
        } else if (registered) sect = 'B2B';
        else if (d.pos && d.pos !== home && d.pos !== '96') sect = 'B2CL?';
        else sect = 'B2CS';
      } else {
        if (registered) sect = 'CDNR';
        else {
          var ur = norm(d.urType);
          if (d.expType) sect = 'CDNUR';
          else if (ur === 'b2cl' || ur === 'expwp' || ur === 'expwop') sect = 'CDNUR';
          else sect = 'B2CS';
        }
      }

      var exportDoc = sect === 'EXP' || (sect === 'CDNUR' && (d.expType || /^EXP/.test(d.urType)));
      var igst = exportDoc || it.igst || (d.pos && d.pos !== home);
      var noPay = (exportDoc && (d.expType === 'WOPAY' || d.urType === 'EXPWOP')) || it.nopay;
      var sign = d.kind === 'CN' ? -1 : 1;
      var diff = d.diffPct ? d.diffPct / 100 : 1;
      if (d.diffPct && d.diffPct !== 65) issues.add('error', 'diffpct', 'Applicable % of Tax Rate must be blank or 65', ref);

      // --- line processing
      var items = {}; // by rate
      var computedVal = 0;
      var anyTaxableLine = false;
      d.lines.forEach(function (L) {
        var hsn = str(L.hsn).replace(/\.0+$/, '').replace(/\s/g, '');
        if (hsnMap[hsn]) { issues.add('info', 'hsnmapped', 'HSN corrected using your Corrections list', hsn + ' → ' + hsnMap[hsn]); hsn = hsnMap[hsn]; }
        else if (invalidHsn[hsn]) issues.add('error', 'hsninvalid', 'HSN rejected by the GST portal as not valid — enter the correct code under Corrections', hsn);
        var rate = num(L.rate);
        var tx = num(L.txval), cess = num(L.cess), qty = num(L.qty);
        var gt = gstTypeCode(L.gstType);
        if (!gt && rate === 0 && opt.zeroRateAsNil && d.kind === 'INV' && sect !== 'EXP') gt = 'NIL';
        if (RATES.indexOf(rate) < 0) issues.add('error', 'rate', 'GST rate not in the GSTN rate list', ref + ': ' + rate + '%');
        if (tx < 0) issues.add('warn', 'negtx', 'Negative taxable value on a line (use a credit note instead)', ref);
        var tax = noPay || gt ? 0 : tx * rate * diff / 100;
        var amt = { txval: tx, iamt: igst ? tax : 0, camt: igst ? 0 : tax / 2, samt: igst ? 0 : tax / 2, csamt: cess };
        computedVal += tx + tax + cess;
        if (src_tot[d.kind]) addAmt(src_tot[d.kind], amt);

        // HSN summary (all outward supplies, net of notes; advances excluded; Non-GST excluded)
        if (sect !== 'AT' && sect !== 'ATADJ' && gt !== 'NONGST') {
          if (!hsn) issues.add('error', 'hsnmissing', 'HSN/SAC missing — line left out of the HSN summary', ref);
          else {
            if (!/^\d{4}(\d{2}(\d{2})?)?$/.test(hsn)) issues.add('warn', 'hsnlen', 'HSN/SAC should be 4, 6 or 8 digits (6 digits minimum when turnover exceeds ₹5 crore)', ref + ': ' + hsn);
            var svc = isService(hsn);
            var uqc = svc ? '' : uqcFull(L.uqc);
            if (!svc && uqc === null) { issues.add('warn', 'uqcbad', 'Unit (UQC) not in the GSTN list — reported as OTH-OTHERS', ref + ': ' + str(L.uqc)); uqc = 'OTH-OTHERS'; }
            if (!svc && uqc === '') { issues.add('warn', 'uqcmissing', 'Unit (UQC) missing for goods — reported as OTH-OTHERS', ref + ': ' + hsn); uqc = 'OTH-OTHERS'; }
            var bucket = registered ? out.hsn.b2b : out.hsn.b2c;
            var hk = hsn + '|' + uqc + '|' + rate;
            var h = bucket[hk];
            if (!h) h = bucket[hk] = { hsn: hsn, uqc: uqc, rate: rate, qty: 0, amt: newAmt(), descs: {} };
            addAmt(h.amt, amt, sign);
            if (!svc) h.qty += (sign < 0 && !opt.reduceQtyOnCN) ? 0 : sign * qty;
            var ds = str(L.desc);
            if (ds) h.descs[ds] = (h.descs[ds] || 0) + Math.abs(tx);
          }
        }

        if (gt) { // Table 8
          var col = gt === 'NIL' ? 0 : gt === 'EXEMPT' ? 1 : 2;
          var nk = (igst ? 'INTR' : 'INTRA') + (registered ? 'B2B' : 'B2C');
          out.nil[nk][col] += sign * tx;
          return;
        }
        if (tx === 0 && cess === 0) { stats.zeroLines++; return; }
        anyTaxableLine = true;
        var item = items[rate];
        if (!item) item = items[rate] = Object.assign({ rate: rate }, newAmt());
        addAmt(item, amt);
      });

      var val = d.valGiven != null ? d.valGiven : computedVal;
      if (d.valGiven != null && Math.abs(d.valGiven - computedVal) > Math.max(1, 0.001 * Math.abs(computedVal)) && sect !== 'AT' && sect !== 'ATADJ')
        issues.add('warn', 'valdiff', 'Total Transaction Value differs from taxable value + tax + cess by more than ₹1', ref + ' (' + r2(d.valGiven) + ' vs ' + r2(computedVal) + ')');

      var itemList = Object.keys(items).map(function (k) { return items[k]; }).sort(function (a, b) { return a.rate - b.rate; });
      if (!itemList.length) {
        // Fully nil/exempt documents live only in Table 8 + HSN. Zero-value documents are left out of the
        // invoice/note tables: the portal rejects them ("Zero items added or added with zero value", RET191239).
        if (!anyTaxableLine && sect !== 'B2CS' && sect !== 'B2CL?' && sect !== 'AT' && sect !== 'ATADJ' && d.lines.some(function (L) { return !gstTypeCode(L.gstType); }))
          issues.add('warn', 'zerodoc', 'Zero-value document left out of ' + sect + ' — the portal rejects documents with zero value. It is still counted in Table 13 (documents issued).', ref);
        return;
      }

      if (sect === 'B2CL?') sect = (val > opt.b2clLimit) ? 'B2CL' : 'B2CS';

      var rec = { d: d, val: val, items: itemList, it: itCode };
      switch (sect) {
        case 'B2B': out.b2b.push(rec); break;
        case 'B2CL': out.b2cl.push(rec); break;
        case 'EXP': out.exp.push(rec);
          if (!d.portCode || !d.sbNum) issues.add('info', 'sbmissing', 'Export invoice without port code / shipping bill — can be added later by amendment', ref);
          break;
        case 'CDNR': out.cdnr.push(rec); break;
        case 'CDNUR':
          rec.urType = d.expType ? (d.expType === 'WPAY' ? 'EXPWP' : 'EXPWOP') : (d.urType || 'B2CL');
          if (rec.urType === 'B2CL' && d.pos === home) issues.add('error', 'cdnurintra', 'UR Type B2CL is only for inter-state notes — this note is intra-state', ref);
          out.cdnur.push(rec); break;
        case 'B2CS':
          itemList.forEach(function (i) {
            var typ = d.etin ? 'E' : 'OE';
            var k = [d.pos, i.rate, typ, d.etin, d.diffPct || ''].join('|');
            var b = out.b2cs[k];
            if (!b) b = out.b2cs[k] = Object.assign({ pos: d.pos, rate: i.rate, typ: typ, etin: d.etin, diffPct: d.diffPct, inter: d.pos !== home }, newAmt());
            addAmt(b, i, sign);
          });
          break;
        case 'AT': case 'ATADJ':
          var tgt = sect === 'AT' ? out.at : out.atadj;
          itemList.forEach(function (i) {
            var k = [d.pos, i.rate, d.diffPct || ''].join('|');
            var b = tgt[k];
            if (!b) b = tgt[k] = Object.assign({ pos: d.pos, rate: i.rate, diffPct: d.diffPct, inter: d.pos !== home }, newAmt());
            addAmt(b, i, 1);
          });
          break;
      }
    });

    // shared numbers between invoices and notes
    var shared = Object.keys(noteNumbers).filter(function (n) { return invNumbers[n]; });
    if (shared.length) shared.slice(0, 25).forEach(function (n) { issues.add('warn', 'sharednum', 'Same number used for an invoice and a credit/debit note', n); });

    // ---- B2CS negative check
    Object.keys(out.b2cs).forEach(function (k) {
      var b = out.b2cs[k];
      if (r2(b.txval) < 0) issues.add('warn', 'b2csneg', 'B2CS is negative for a place of supply / rate after netting unregistered credit notes. The portal may reject a negative B2CS row — report the excess via B2CSA / adjust in the original period.', posLabel(b.pos) + ' @ ' + b.rate + '%: ' + r2(b.txval));
    });

    // ---- HSN finalise
    if (opt.mergeNoteUnits) { mergeUnits(out.hsn.b2b, issues); mergeUnits(out.hsn.b2c, issues); }
    var hsnOut = { b2b: finaliseHsn(out.hsn.b2b, issues, 'B2B'), b2c: finaliseHsn(out.hsn.b2c, issues, 'B2C') };

    // ---- Docs (Table 13)
    var docRows = finaliseSeries(out.series, opt.gapsAsCancelled);

    var b2csList = Object.keys(out.b2cs).map(function (k) { return out.b2cs[k]; })
      .filter(function (b) { return r2(b.txval) !== 0 || r2(b.csamt) !== 0; })
      .sort(function (a, b) { return a.pos.localeCompare(b.pos) || a.rate - b.rate; });
    var atList = Object.keys(out.at).map(function (k) { return out.at[k]; });
    var atadjList = Object.keys(out.atadj).map(function (k) { return out.atadj[k]; });

    var res = {
      options: opt, home: home, period: period, detectedPeriod: detectedPeriod, source: { sheet: src.sheet, headerRow: src.headerRow, columns: src.columns },
      stats: stats, sourceTotals: src_tot,
      b2b: out.b2b, b2cl: out.b2cl, exp: out.exp, cdnr: out.cdnr, cdnur: out.cdnur,
      b2cs: b2csList, at: atList, atadj: atadjList, nil: out.nil,
      hsn: hsnOut, docs: docRows, issues: null
    };
    res.summary = summarise(res);
    reconcile(res, issues);
    res.issues = issues.sorted();
    return res;
  }

  /* One row per HSN + rate. Rows of the same HSN and rate booked in different units (e.g. sales in
     PCS, credit notes in NOS) are folded into the main unit — the one carrying the largest value —
     because the offline tool / portal reject an HSN repeated with several UQCs. */
  function mergeUnits(map, issues) {
    var groups = {};
    Object.keys(map).forEach(function (k) { var h = map[k]; var g = h.hsn + '|' + h.rate; (groups[g] = groups[g] || []).push(k); });
    Object.keys(groups).forEach(function (g) {
      var keys = groups[g];
      if (keys.length < 2) return;
      var main = null;
      keys.forEach(function (k) {
        if (!main) { main = k; return; }
        var a = map[k].amt.txval, b = map[main].amt.txval;
        if ((a > 0 && b <= 0) || ((a > 0) === (b > 0) && Math.abs(a) > Math.abs(b))) main = k;
      });
      var m = map[main];
      keys.forEach(function (k) {
        if (k === main) return;
        var h = map[k];
        addAmt(m.amt, h.amt); m.qty += h.qty;
        for (var ds in h.descs) m.descs[ds] = (m.descs[ds] || 0) + h.descs[ds];
        issues.add('info', 'hsnmerge', 'Same HSN and rate booked in more than one unit — quantities merged into one UQC (one row per HSN + rate)', h.hsn + ' @ ' + h.rate + '%: ' + (h.uqc || 'NA') + ' → ' + (m.uqc || 'NA'));
        delete map[k];
      });
    });
  }

  function finaliseHsn(map, issues, label) {
    return Object.keys(map).map(function (k) {
      var h = map[k];
      var desc = '';
      var best = -1;
      for (var ds in h.descs) if (h.descs[ds] > best) { best = h.descs[ds]; desc = ds; }
      var row = { hsn: h.hsn, desc: desc.replace(/\s+/g, ' ').trim(), uqc: h.uqc, qty: Math.round(h.qty * 100) / 100, rate: h.rate,
        txval: r2(h.amt.txval), iamt: r2(h.amt.iamt), camt: r2(h.amt.camt), samt: r2(h.amt.samt), csamt: r2(h.amt.csamt) };
      if (row.txval < 0) issues.add('warn', 'hsnneg', 'HSN ' + label + ' row is negative — credit notes for this HSN/rate exceed this month\'s sales (typically returns of earlier months). Check the offline tool accepts it, or net it against another period.', h.hsn + ' @ ' + h.rate + '%: ' + row.txval);
      if (row.qty < 0) {
        issues.add('info', 'hsnnegqty', 'HSN ' + label + ' net quantity was negative after credit notes — reported as 0 (the offline tool rejects negative quantity)', h.hsn + ' ' + row.uqc + ': ' + row.qty);
        row.qty = 0;
      }
      return row;
    }).filter(function (r) { return r.txval !== 0 || r.qty !== 0 || r.csamt !== 0; })
      .sort(function (a, b) { return a.hsn.localeCompare(b.hsn) || a.rate - b.rate || a.uqc.localeCompare(b.uqc); });
  }

  /* Table 13: split each document number into prefix + running number + suffix. */
  function registerSeries(series, nature, no, cancelled) {
    var m = /^(.*?)(\d+)(\D*)$/.exec(no);
    var key, n = null, len = 0;
    if (m) { key = nature + '|' + m[1] + '|' + m[2].length + '|' + m[3]; n = BigInt(m[2]); len = m[2].length; }
    else key = nature + '|' + no + '|x';
    var s = series[key];
    if (!s) s = series[key] = { nature: nature, prefix: m ? m[1] : no, suffix: m ? m[3] : '', len: len, nums: {}, cancelled: 0, min: null, max: null, minNo: no, maxNo: no, count: 0 };
    if (s.nums[no]) return;
    s.nums[no] = 1; s.count++;
    if (cancelled) s.cancelled++;
    if (n !== null) {
      if (s.min === null || n < s.min) { s.min = n; s.minNo = no; }
      if (s.max === null || n > s.max) { s.max = n; s.maxNo = no; }
    }
  }
  function finaliseSeries(series, gapsAsCancelled) {
    var natureOrder = { INV: 1, RV: 2, DN: 3, CN: 4 };
    return Object.keys(series).map(function (k) {
      var s = series[k];
      var range = s.min !== null ? Number(s.max - s.min + 1n) : s.count;
      var gaps = Math.max(0, range - s.count);
      var total = gapsAsCancelled ? range : s.count;
      var cancel = s.cancelled + (gapsAsCancelled ? gaps : 0);
      return { nature: s.nature, natureLabel: DOC_NATURE[s.nature].typ, docNum: DOC_NATURE[s.nature].num,
        from: s.minNo, to: s.maxNo, total: total, cancel: cancel, net: total - cancel, gaps: gaps, issued: s.count };
    }).sort(function (a, b) { return natureOrder[a.nature] - natureOrder[b.nature] || a.from.localeCompare(b.from); });
  }

  function sumItems(list, sign) {
    var a = newAmt();
    list.forEach(function (r) { r.items.forEach(function (i) { addAmt(a, i, sign && sign(r)); }); });
    return a;
  }
  function sumFlat(list) { var a = newAmt(); list.forEach(function (b) { addAmt(a, b); }); return a; }
  function round(a) { return { txval: r2(a.txval), iamt: r2(a.iamt), camt: r2(a.camt), samt: r2(a.samt), csamt: r2(a.csamt) }; }

  function summarise(res) {
    var noteSign = function (r) { return r.d.kind === 'CN' ? -1 : 1; };
    var nilTot = 0;
    Object.keys(res.nil).forEach(function (k) { nilTot += res.nil[k][0] + res.nil[k][1] + res.nil[k][2]; });
    var rows = [
      { key: 'b2b', label: 'B2B, SEZ, DE — Table 4A, 4B, 6B, 6C', sheet: 'b2b,sez,de', docs: res.b2b.length, parties: uniq(res.b2b.map(function (r) { return r.d.gstin; })), amt: round(sumItems(res.b2b)), value: r2(res.b2b.reduce(function (s, r) { return s + r.val; }, 0)) },
      { key: 'b2cl', label: 'B2C Large — Table 5', sheet: 'b2cl', docs: res.b2cl.length, amt: round(sumItems(res.b2cl)), value: r2(res.b2cl.reduce(function (s, r) { return s + r.val; }, 0)) },
      { key: 'exp', label: 'Exports — Table 6A', sheet: 'exp', docs: res.exp.length, amt: round(sumItems(res.exp)), value: r2(res.exp.reduce(function (s, r) { return s + r.val; }, 0)) },
      { key: 'b2cs', label: 'B2C Others (net of unregistered notes) — Table 7', sheet: 'b2cs', docs: res.b2cs.length, rowsLabel: 'rows', amt: round(sumFlat(res.b2cs)) },
      { key: 'nil', label: 'Nil rated / Exempt / Non-GST — Table 8', sheet: 'exemp', docs: null, amt: { txval: r2(nilTot), iamt: 0, camt: 0, samt: 0, csamt: 0 } },
      { key: 'cdnr', label: 'Credit / Debit Notes (registered) — Table 9B', sheet: 'cdnr', docs: res.cdnr.length, parties: uniq(res.cdnr.map(function (r) { return r.d.gstin; })), amt: round(sumItems(res.cdnr, noteSign)), signed: true },
      { key: 'cdnur', label: 'Credit / Debit Notes (unregistered) — Table 9B', sheet: 'cdnur', docs: res.cdnur.length, amt: round(sumItems(res.cdnur, noteSign)), signed: true },
      { key: 'at', label: 'Advances received — Table 11A', sheet: 'at', docs: res.at.length, rowsLabel: 'rows', amt: round(sumFlat(res.at)) },
      { key: 'atadj', label: 'Advances adjusted — Table 11B', sheet: 'atadj', docs: res.atadj.length, rowsLabel: 'rows', amt: round(sumFlat(res.atadj)), signed: true, negate: true }
    ];
    var net = newAmt();
    rows.forEach(function (r) { addAmt(net, r.amt, r.negate ? -1 : 1); });
    var hb = newAmt(), hc = newAmt();
    res.hsn.b2b.forEach(function (h) { addAmt(hb, h); });
    res.hsn.b2c.forEach(function (h) { addAmt(hc, h); });
    return { rows: rows, net: round(net), hsnB2B: round(hb), hsnB2C: round(hc),
      docsTotal: res.docs.reduce(function (s, d) { return s + d.total; }, 0), docsCancelled: res.docs.reduce(function (s, d) { return s + d.cancel; }, 0) };
  }
  function uniq(a) { var o = {}; a.forEach(function (x) { o[x] = 1; }); return Object.keys(o).length; }

  function reconcile(res, issues) {
    var S = res.summary, by = {};
    S.rows.forEach(function (r) { by[r.key] = r.amt; });
    var nilReg = 0, nilUnreg = 0;
    ['INTRB2B', 'INTRAB2B'].forEach(function (k) { nilReg += res.nil[k][0] + res.nil[k][1]; });
    ['INTRB2C', 'INTRAB2C'].forEach(function (k) { nilUnreg += res.nil[k][0] + res.nil[k][1]; });
    var expB2B = r2(by.b2b.txval + by.cdnr.txval + nilReg);
    var expB2C = r2(by.b2cl.txval + by.exp.txval + by.b2cs.txval + by.cdnur.txval + nilUnreg);
    S.recon = [
      { label: 'HSN (B2B) vs B2B + CDNR + nil/exempt to registered', expected: expB2B, hsn: S.hsnB2B.txval, diff: r2(S.hsnB2B.txval - expB2B) },
      { label: 'HSN (B2C) vs B2CL + EXP + B2CS + CDNUR + nil/exempt to unregistered', expected: expB2C, hsn: S.hsnB2C.txval, diff: r2(S.hsnB2C.txval - expB2C) }
    ];
    var tb = r2(S.hsnB2B.iamt + S.hsnB2B.camt + S.hsnB2B.samt), tbx = r2(by.b2b.iamt + by.b2b.camt + by.b2b.samt + by.cdnr.iamt + by.cdnr.camt + by.cdnr.samt);
    S.recon.push({ label: 'Tax in HSN (B2B) vs tax in B2B + CDNR', expected: tbx, hsn: tb, diff: r2(tb - tbx) });
    var tc = r2(S.hsnB2C.iamt + S.hsnB2C.camt + S.hsnB2C.samt), tcx = r2(by.b2cl.iamt + by.exp.iamt + by.b2cs.iamt + by.b2cs.camt + by.b2cs.samt + by.cdnur.iamt + by.cdnur.camt + by.cdnur.samt);
    S.recon.push({ label: 'Tax in HSN (B2C) vs tax in B2CL + EXP + B2CS + CDNUR', expected: tcx, hsn: tc, diff: r2(tc - tcx) });
    S.recon.forEach(function (r) {
      if (Math.abs(r.diff) > 1) issues.add('warn', 'recon', 'HSN summary does not reconcile to the sections (lines without HSN or rounding)', r.label + ': ' + r.diff);
    });
  }

  /* ------------------------------------------------------------------ *
   * 6. GSTR-1 JSON
   * ------------------------------------------------------------------ */

  function itmDet(i, inter, withNum, n) {
    var det = { txval: r2(i.txval), rt: i.rate };
    if (inter) det.iamt = r2(i.iamt); else { det.camt = r2(i.camt); det.samt = r2(i.samt); }
    det.csamt = r2(i.csamt);
    return withNum ? { num: n, itm_det: det } : det;
  }
  function isInter(rec, home) {
    var it = INV_TYPES[rec.it] || INV_TYPES.R;
    return it.igst || rec.d.pos !== home;
  }

  function buildJson(res) {
    var opt = res.options, home = res.home;
    var j = { gstin: str(opt.supplierGstin).toUpperCase(), fp: res.period, version: opt.jsonVersion || 'GST3.2.3', hash: 'hash' };

    function partyGroup(list, mk) {
      var by = {}, order = [];
      list.forEach(function (r) { var g = r.d.gstin; if (!by[g]) { by[g] = []; order.push(g); } by[g].push(mk(r)); });
      return order.map(function (g) { return { ctin: g, list: by[g] }; });
    }
    if (res.b2b.length) j.b2b = partyGroup(res.b2b, function (r) {
      var inter = isInter(r, home);
      var it = INV_TYPES[r.it];
      var o = { inum: r.d.no, idt: dmy(r.d.date), val: r2(r.val), pos: r.d.pos, rchrg: r.d.rchrg, inv_typ: it.json || r.it };
      if (r.d.etin) o.etin = r.d.etin;
      if (r.d.diffPct) o.diff_percent = r.d.diffPct / 100;
      o.itms = r.items.map(function (i, n) { return itmDet(i, inter, true, n + 1); });
      return o;
    }).map(function (g) { return { ctin: g.ctin, inv: g.list }; });

    if (res.b2cl.length) {
      var byPos = {}, posOrder = [];
      res.b2cl.forEach(function (r) {
        if (!byPos[r.d.pos]) { byPos[r.d.pos] = []; posOrder.push(r.d.pos); }
        var o = { inum: r.d.no, idt: dmy(r.d.date), val: r2(r.val) };
        if (r.d.etin) o.etin = r.d.etin;
        if (r.d.diffPct) o.diff_percent = r.d.diffPct / 100;
        o.itms = r.items.map(function (i, n) { return { num: n + 1, itm_det: { rt: i.rate, txval: r2(i.txval), iamt: r2(i.iamt), csamt: r2(i.csamt) } }; });
        byPos[r.d.pos].push(o);
      });
      j.b2cl = posOrder.map(function (p) { return { pos: p, inv: byPos[p] }; });
    }

    if (res.b2cs.length) j.b2cs = res.b2cs.map(function (b) {
      var o = { sply_ty: b.inter ? 'INTER' : 'INTRA', pos: b.pos, typ: b.typ, rt: b.rate, txval: r2(b.txval) };
      if (b.etin) o.etin = b.etin;
      if (b.diffPct) o.diff_percent = b.diffPct / 100;
      if (b.inter) o.iamt = r2(b.iamt); else { o.camt = r2(b.camt); o.samt = r2(b.samt); }
      o.csamt = r2(b.csamt);
      return o;
    });

    if (res.exp.length) {
      var byT = { WPAY: [], WOPAY: [] };
      res.exp.forEach(function (r) {
        var o = { inum: r.d.no, idt: dmy(r.d.date), val: r2(r.val) };
        if (r.d.portCode) o.sbpcode = r.d.portCode;
        if (r.d.sbNum) o.sbnum = r.d.sbNum;
        if (r.d.sbDate) o.sbdt = dmy(r.d.sbDate);
        if (r.d.diffPct) o.diff_percent = r.d.diffPct / 100;
        o.itms = r.items.map(function (i) { return { txval: r2(i.txval), rt: i.rate, iamt: r2(i.iamt), csamt: r2(i.csamt) }; });
        byT[r.d.expType === 'WPAY' ? 'WPAY' : 'WOPAY'].push(o);
      });
      j.exp = ['WPAY', 'WOPAY'].filter(function (t) { return byT[t].length; }).map(function (t) { return { exp_typ: t, inv: byT[t] }; });
    }

    if (res.cdnr.length) j.cdnr = partyGroup(res.cdnr, function (r) {
      var inter = isInter(r, home);
      var it = INV_TYPES[r.it];
      var o = { ntty: r.d.kind === 'CN' ? 'C' : 'D', nt_num: r.d.no, nt_dt: dmy(r.d.date), val: r2(r.val), pos: r.d.pos, rchrg: r.d.rchrg, inv_typ: it.json || r.it };
      if (r.d.diffPct) o.diff_percent = r.d.diffPct / 100;
      o.itms = r.items.map(function (i, n) { return itmDet(i, inter, true, n + 1); });
      return o;
    }).map(function (g) { return { ctin: g.ctin, nt: g.list }; });

    if (res.cdnur.length) j.cdnur = res.cdnur.map(function (r) {
      var o = { typ: r.urType, ntty: r.d.kind === 'CN' ? 'C' : 'D', nt_num: r.d.no, nt_dt: dmy(r.d.date), val: r2(r.val) };
      if (r.urType === 'B2CL') o.pos = r.d.pos;
      if (r.d.diffPct) o.diff_percent = r.d.diffPct / 100;
      o.itms = r.items.map(function (i, n) { return { num: n + 1, itm_det: { rt: i.rate, txval: r2(i.txval), iamt: r2(i.iamt), csamt: r2(i.csamt) } }; });
      return o;
    });

    var nilRows = EXEMP_ROWS.map(function (e) { var v = res.nil[e[0]]; return { sply_ty: e[0], nil_amt: r2(v[0]), expt_amt: r2(v[1]), ngsup_amt: r2(v[2]) }; });
    if (nilRows.some(function (n) { return n.nil_amt || n.expt_amt || n.ngsup_amt; })) j.nil = { inv: nilRows };

    function advJson(list, amtKey) {
      var by = {}, order = [];
      list.forEach(function (b) {
        var k = b.pos + '|' + (b.diffPct || '');
        if (!by[k]) { by[k] = { pos: b.pos, sply_ty: b.inter ? 'INTER' : 'INTRA', itms: [] }; if (b.diffPct) by[k].diff_percent = b.diffPct / 100; order.push(k); }
        var it = { rt: b.rate }; it[amtKey] = r2(b.txval);
        if (b.inter) it.iamt = r2(b.iamt); else { it.camt = r2(b.camt); it.samt = r2(b.samt); }
        it.csamt = r2(b.csamt);
        by[k].itms.push(it);
      });
      return order.map(function (k) { return by[k]; });
    }
    if (res.at.length) j.at = advJson(res.at, 'ad_amt');
    if (res.atadj.length) j.txpd = advJson(res.atadj, 'ad_amt');

    function hsnJson(list) {
      return list.map(function (h, n) {
        var o = { num: n + 1, hsn_sc: h.hsn };
        if (opt.hsnDesc && h.desc) o.desc = h.desc.slice(0, 30);
        o.uqc = h.uqc ? h.uqc.split('-')[0] : 'NA';
        o.qty = h.uqc ? h.qty : 0;
        o.rt = h.rate; o.txval = h.txval; o.iamt = h.iamt; o.camt = h.camt; o.samt = h.samt; o.csamt = h.csamt;
        return o;
      });
    }
    if (res.hsn.b2b.length || res.hsn.b2c.length) {
      j.hsn = {};
      if (res.hsn.b2b.length) j.hsn.hsn_b2b = hsnJson(res.hsn.b2b);
      if (res.hsn.b2c.length) j.hsn.hsn_b2c = hsnJson(res.hsn.b2c);
    }

    if (res.docs.length) {
      var byN = {}, nOrder = [];
      res.docs.forEach(function (d) {
        if (!byN[d.docNum]) { byN[d.docNum] = { doc_num: d.docNum, doc_typ: d.natureLabel, docs: [] }; nOrder.push(d.docNum); }
        byN[d.docNum].docs.push({ num: byN[d.docNum].docs.length + 1, from: d.from, to: d.to, totnum: d.total, cancel: d.cancel, net_issue: d.net });
      });
      j.doc_issue = { doc_det: nOrder.map(function (n) { return byN[n]; }) };
    }
    return j;
  }

  /* Split a large return into several JSON files (GSTN upload limit ~5 MB). Summary
     sections go in part 1; invoice-level sections are spread across parts by recipient. */
  function splitJson(j, maxBytes) {
    maxBytes = maxBytes || 4800000;
    var whole = JSON.stringify(j);
    if (whole.length <= maxBytes) return [j];
    var head = { gstin: j.gstin, fp: j.fp, version: j.version, hash: j.hash };
    var parts = [], cur = Object.assign({}, head), curSize = JSON.stringify(cur).length;
    ['b2cs', 'nil', 'hsn', 'doc_issue', 'at', 'txpd'].forEach(function (k) { if (j[k]) { cur[k] = j[k]; } });
    curSize = JSON.stringify(cur).length;
    function push() { parts.push(cur); cur = Object.assign({}, head); curSize = JSON.stringify(cur).length; }
    ['b2b', 'cdnr', 'b2cl', 'exp', 'cdnur'].forEach(function (k) {
      if (!j[k]) return;
      j[k].forEach(function (grp) {
        var sz = JSON.stringify(grp).length + 1;
        if (curSize + sz + 20 > maxBytes && curSize > 200) push();
        (cur[k] = cur[k] || []).push(grp);
        curSize += sz + (cur[k].length === 1 ? k.length + 6 : 0);
      });
    });
    push();
    return parts;
  }

  /* ------------------------------------------------------------------ *
   * 7. Excel writer (GSTR1_Excel_Workbook_Template V2.1 layout)
   * ------------------------------------------------------------------ */

  function xmlEsc(s) {
    return String(s).replace(/[&<>"]/g, function (c) { return c === '&' ? '&amp;' : c === '<' ? '&lt;' : c === '>' ? '&gt;' : '&quot;'; })
      .replace(/[\u0000-\u0008\u000B\u000C\u000E-\u001F]/g, '');
  }
  function colName(i) { var s = ''; i++; while (i > 0) { var m = (i - 1) % 26; s = String.fromCharCode(65 + m) + s; i = Math.floor((i - 1) / 26); } return s; }

  // style ids (see STYLES_XML)
  var ST = { def: 0, title: 1, help: 2, sumLabel: 3, sumVal: 4, head: 5, text: 6, num2: 7, date: 8, int: 9, sumNum: 10 };
  var STYLES_XML = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n' +
    '<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">' +
    '<numFmts count="1"><numFmt numFmtId="164" formatCode="dd\\-mmm\\-yyyy"/></numFmts>' +
    '<fonts count="4"><font><sz val="11"/><name val="Calibri"/><family val="2"/></font>' +
    '<font><b/><sz val="11"/><color rgb="FFFFFFFF"/><name val="Calibri"/><family val="2"/></font>' +
    '<font><b/><sz val="11"/><name val="Calibri"/><family val="2"/></font>' +
    '<font><b/><u/><sz val="11"/><color rgb="FFFFFFFF"/><name val="Calibri"/><family val="2"/></font></fonts>' +
    '<fills count="4"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill>' +
    '<fill><patternFill patternType="solid"><fgColor rgb="FF0070C0"/><bgColor indexed="64"/></patternFill></fill>' +
    '<fill><patternFill patternType="solid"><fgColor rgb="FFF8CBAD"/><bgColor indexed="64"/></patternFill></fill></fills>' +
    '<borders count="2"><border><left/><right/><top/><bottom/><diagonal/></border>' +
    '<border><left style="thin"><color auto="1"/></left><right style="thin"><color auto="1"/></right><top style="thin"><color auto="1"/></top><bottom style="thin"><color auto="1"/></bottom><diagonal/></border></borders>' +
    '<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>' +
    '<cellXfs count="11">' +
    '<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>' +                                           // 0 default
    '<xf numFmtId="0" fontId="1" fillId="2" borderId="0" xfId="0" applyFont="1" applyFill="1"/>' +               // 1 title
    '<xf numFmtId="0" fontId="3" fillId="2" borderId="0" xfId="0" applyFont="1" applyFill="1"/>' +               // 2 help
    '<xf numFmtId="0" fontId="1" fillId="2" borderId="0" xfId="0" applyFont="1" applyFill="1" applyAlignment="1"><alignment wrapText="1"/></xf>' + // 3 summary label
    '<xf numFmtId="0" fontId="2" fillId="0" borderId="0" xfId="0" applyFont="1"/>' +                             // 4 summary value
    '<xf numFmtId="0" fontId="2" fillId="3" borderId="1" xfId="0" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1"><alignment horizontal="center" vertical="center" wrapText="1"/></xf>' + // 5 header
    '<xf numFmtId="49" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/>' +                     // 6 text
    '<xf numFmtId="2" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/>' +                      // 7 0.00
    '<xf numFmtId="164" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/>' +                    // 8 date
    '<xf numFmtId="1" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/>' +                      // 9 integer
    '<xf numFmtId="2" fontId="2" fillId="0" borderId="0" xfId="0" applyFont="1" applyNumberFormat="1"/>' +        // 10 summary number
    '</cellXfs><cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles></styleSheet>';

  function SST() { this.idx = {}; this.list = []; this.count = 0; }
  SST.prototype.id = function (s) {
    this.count++;
    var i = this.idx['$' + s];
    if (i === undefined) { i = this.idx['$' + s] = this.list.length; this.list.push(s); }
    return i;
  };
  SST.prototype.xml = function () {
    var parts = ['<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="' + this.count + '" uniqueCount="' + this.list.length + '">'];
    for (var i = 0; i < this.list.length; i++) {
      var s = this.list[i];
      parts.push('<si><t' + (/^\s|\s$/.test(s) ? ' xml:space="preserve"' : '') + '>' + xmlEsc(s) + '</t></si>');
    }
    parts.push('</sst>');
    return parts.join('');
  };

  // A cell value: string -> shared string; number; {d:{y,m,d}} -> date; {n:number,s:style}
  function cellXml(sst, ref, v, style) {
    if (v === null || v === undefined || v === '') return '';
    var s = style != null ? ' s="' + style + '"' : '';
    if (typeof v === 'number') return isFinite(v) ? '<c r="' + ref + '"' + s + '><v>' + v + '</v></c>' : '';
    if (v && v.y) return '<c r="' + ref + '" s="' + ST.date + '"><v>' + dateSerial(v) + '</v></c>';
    return '<c r="' + ref + '"' + s + ' t="s"><v>' + sst.id(String(v)) + '</v></c>';
  }

  function sheetXml(sst, spec, sum, rows, colStyles) {
    var ncol = spec.h.length;
    var out = ['<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">'];
    out.push('<sheetViews><sheetView workbookViewId="0"><pane ySplit="4" topLeftCell="A5" activePane="bottomLeft" state="frozen"/></sheetView></sheetViews>');
    out.push('<sheetFormatPr defaultRowHeight="15"/><cols>');
    for (var c = 0; c < ncol; c++) out.push('<col min="' + (c + 1) + '" max="' + (c + 1) + '" width="' + (spec.w[c] || 14) + '" customWidth="1"/>');
    out.push('</cols><sheetData>');
    // row 1: title + HELP
    var r1 = '<row r="1">';
    for (c = 0; c < spec.r1.length; c++) {
      if (spec.r1[c]) r1 += cellXml(sst, colName(c) + '1', spec.r1[c], spec.r1[c] === 'HELP' ? ST.help : ST.title);
      else r1 += '<c r="' + colName(c) + '1" s="' + ST.title + '"/>';
    }
    out.push(r1 + '</row>');
    // row 2: summary labels
    var r2s = '<row r="2">';
    for (c = 0; c < ncol; c++) {
      var lab = spec.r2[c];
      r2s += lab ? cellXml(sst, colName(c) + '2', lab, ST.sumLabel) : '<c r="' + colName(c) + '2" s="' + ST.sumLabel + '"/>';
    }
    out.push(r2s + '</row>');
    // row 3: summary values
    var r3 = '<row r="3">';
    for (c = 0; c < ncol; c++) {
      var sv = sum && sum[c];
      if (sv !== undefined && sv !== null) r3 += cellXml(sst, colName(c) + '3', sv, Number.isInteger(sv) ? ST.sumVal : ST.sumNum);
    }
    out.push(r3 + '</row>');
    // row 4: headers
    var r4 = '<row r="4" ht="45" customHeight="1">';
    for (c = 0; c < ncol; c++) r4 += cellXml(sst, colName(c) + '4', spec.h[c], ST.head);
    out.push(r4 + '</row>');
    var cn = []; for (c = 0; c < ncol; c++) cn.push(colName(c));
    for (var i = 0; i < rows.length; i++) {
      var rn = i + 5, row = rows[i], x = '<row r="' + rn + '">';
      for (c = 0; c < ncol; c++) x += cellXml(sst, cn[c] + rn, row[c], colStyles ? colStyles[c] : null);
      out.push(x + '</row>');
    }
    out.push('</sheetData></worksheet>');
    return out.join('');
  }

  function excelRows(res) {
    var home = res.home, R = {}, S = {};
    function pd(r) { return r.d.diffPct ? r.d.diffPct : ''; }
    function sumCol(rows, c) { return r2(rows.reduce(function (s, r) { return s + (typeof r[c] === 'number' ? r[c] : 0); }, 0)); }
    function distinct(rows, c) { var o = {}; rows.forEach(function (r) { if (r[c] !== '' && r[c] != null) o[r[c]] = 1; }); return Object.keys(o).length; }
    function distinctVal(rows, kc, vc) { var o = {}; rows.forEach(function (r) { o[r[kc]] = r[vc]; }); return r2(Object.keys(o).reduce(function (s, k) { return s + o[k]; }, 0)); }

    var b2b = [];
    res.b2b.forEach(function (r) {
      r.items.forEach(function (i) {
        b2b.push([r.d.gstin, r.d.name, r.d.no, r.d.date, r2(r.val), posLabel(r.d.pos), r.d.rchrg, pd(r), INV_TYPES[r.it].label, r.d.etin, i.rate, r2(i.txval), r2(i.csamt)]);
      });
    });
    R['b2b,sez,de'] = b2b; S['b2b,sez,de'] = [distinct(b2b, 0), null, distinct(b2b, 2), null, distinctVal(b2b, 2, 4), null, null, null, null, null, null, sumCol(b2b, 11), sumCol(b2b, 12)];

    var b2cl = [];
    res.b2cl.forEach(function (r) { r.items.forEach(function (i) { b2cl.push([r.d.no, r.d.date, r2(r.val), posLabel(r.d.pos), pd(r), i.rate, r2(i.txval), r2(i.csamt), r.d.etin]); }); });
    R.b2cl = b2cl; S.b2cl = [distinct(b2cl, 0), null, distinctVal(b2cl, 0, 2), null, null, null, sumCol(b2cl, 6), sumCol(b2cl, 7)];

    var b2cs = res.b2cs.map(function (b) { return [b.typ, posLabel(b.pos), b.diffPct || '', b.rate, r2(b.txval), r2(b.csamt), b.etin]; });
    R.b2cs = b2cs; S.b2cs = [null, null, null, null, sumCol(b2cs, 4), sumCol(b2cs, 5)];

    var cdnr = [];
    res.cdnr.forEach(function (r) {
      r.items.forEach(function (i) {
        cdnr.push([r.d.gstin, r.d.name, r.d.no, r.d.date, r.d.kind === 'CN' ? 'C' : 'D', posLabel(r.d.pos), r.d.rchrg, INV_TYPES[r.it].label, r2(r.val), pd(r), i.rate, r2(i.txval), r2(i.csamt)]);
      });
    });
    R.cdnr = cdnr; S.cdnr = [distinct(cdnr, 0), null, distinct(cdnr, 2), null, null, null, null, null, distinctVal(cdnr, 2, 8), null, null, sumCol(cdnr, 11), sumCol(cdnr, 12)];

    var cdnur = [];
    res.cdnur.forEach(function (r) {
      r.items.forEach(function (i) { cdnur.push([r.urType, r.d.no, r.d.date, r.d.kind === 'CN' ? 'C' : 'D', r.urType === 'B2CL' ? posLabel(r.d.pos) : '', r2(r.val), pd(r), i.rate, r2(i.txval), r2(i.csamt)]); });
    });
    R.cdnur = cdnur; S.cdnur = [null, distinct(cdnur, 1), null, null, null, distinctVal(cdnur, 1, 5), null, null, sumCol(cdnur, 8), sumCol(cdnur, 9)];

    var exp = [];
    res.exp.forEach(function (r) {
      r.items.forEach(function (i) { exp.push([r.d.expType, r.d.no, r.d.date, r2(r.val), r.d.portCode, r.d.sbNum, r.d.sbDate || '', i.rate, r2(i.txval), r2(i.csamt)]); });
    });
    R.exp = exp; S.exp = [null, distinct(exp, 1), null, distinctVal(exp, 1, 3), null, distinct(exp, 5), null, null, null, sumCol(exp, 8)];

    var at = res.at.map(function (b) { return [posLabel(b.pos), b.diffPct || '', b.rate, r2(b.txval), r2(b.csamt)]; });
    R.at = at; S.at = [null, null, null, sumCol(at, 3), sumCol(at, 4)];
    var atadj = res.atadj.map(function (b) { return [posLabel(b.pos), b.diffPct || '', b.rate, r2(b.txval), r2(b.csamt)]; });
    R.atadj = atadj; S.atadj = [null, null, null, sumCol(atadj, 3), sumCol(atadj, 4)];

    var exemp = EXEMP_ROWS.map(function (e) { var v = res.nil[e[0]]; return [e[1], r2(v[0]), r2(v[1]), r2(v[2])]; });
    R.exemp = exemp; S.exemp = [null, sumCol(exemp, 1), sumCol(exemp, 2), sumCol(exemp, 3)];

    function hsnRows(list) {
      return list.map(function (h) {
        return [h.hsn, res.options.hsnDesc ? h.desc.slice(0, 30) : '', h.uqc, h.uqc ? h.qty : '', r2(h.txval + h.iamt + h.camt + h.samt + h.csamt), h.rate, h.txval, h.iamt, h.camt, h.samt, h.csamt];
      });
    }
    var hb = hsnRows(res.hsn.b2b), hc = hsnRows(res.hsn.b2c);
    R['hsn(b2b)'] = hb; R['hsn(b2c)'] = hc;
    [['hsn(b2b)', hb], ['hsn(b2c)', hc]].forEach(function (p) {
      S[p[0]] = [distinct(p[1], 0), null, null, null, sumCol(p[1], 4), null, sumCol(p[1], 6), sumCol(p[1], 7), sumCol(p[1], 8), sumCol(p[1], 9), sumCol(p[1], 10)];
    });

    var docs = res.docs.map(function (d) { return [d.natureLabel, d.from, d.to, d.total, d.cancel]; });
    R.docs = docs; S.docs = [null, null, null, docs.reduce(function (s, r) { return s + r[3]; }, 0), docs.reduce(function (s, r) { return s + r[4]; }, 0)];
    return { rows: R, sums: S };
  }

  // column styles per sheet (text for ids / HSN so leading zeros survive)
  var COL_STYLES = {
    'b2b,sez,de': [6, 6, 6, null, 7, 6, 6, 7, 6, 6, 7, 7, 7],
    b2cl: [6, null, 7, 6, 7, 7, 7, 7, 6],
    b2cs: [6, 6, 7, 7, 7, 7, 6],
    cdnr: [6, 6, 6, null, 6, 6, 6, 6, 7, 7, 7, 7, 7],
    cdnur: [6, 6, null, 6, 6, 7, 7, 7, 7, 7],
    exp: [6, 6, null, 7, 6, 6, null, 7, 7, 7],
    at: [6, 7, 7, 7, 7], atadj: [6, 7, 7, 7, 7],
    exemp: [6, 7, 7, 7],
    'hsn(b2b)': [6, 6, 6, 7, 7, 7, 7, 7, 7, 7, 7],
    'hsn(b2c)': [6, 6, 6, 7, 7, 7, 7, 7, 7, 7, 7],
    docs: [6, 6, 6, 9, 9]
  };

  function buildExcel(res) {
    var sst = new SST();
    var data = excelRows(res);
    var files = {};
    var sheetsXml = '', relsXml = '', ctSheets = '';
    TEMPLATE.sheets.forEach(function (spec, i) {
      var n = i + 1;
      var rows = data.rows[spec.n] || [];
      var sum = data.sums[spec.n] || null;
      files['xl/worksheets/sheet' + n + '.xml'] = fflate.strToU8(sheetXml(sst, spec, sum, rows, COL_STYLES[spec.n]));
      sheetsXml += '<sheet name="' + xmlEsc(spec.n) + '" sheetId="' + n + '" r:id="rId' + n + '"/>';
      relsXml += '<Relationship Id="rId' + n + '" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet' + n + '.xml"/>';
      ctSheets += '<Override PartName="/xl/worksheets/sheet' + n + '.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>';
    });
    var k = TEMPLATE.sheets.length;
    relsXml += '<Relationship Id="rId' + (k + 1) + '" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>';
    relsXml += '<Relationship Id="rId' + (k + 2) + '" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings" Target="sharedStrings.xml"/>';
    files['[Content_Types].xml'] = fflate.strToU8('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">' +
      '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/>' +
      '<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>' + ctSheets +
      '<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>' +
      '<Override PartName="/xl/sharedStrings.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"/>' +
      '<Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>' +
      '<Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/></Types>');
    files['_rels/.rels'] = fflate.strToU8('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
      '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>' +
      '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/>' +
      '<Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties" Target="docProps/app.xml"/></Relationships>');
    files['docProps/core.xml'] = fflate.strToU8('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">' +
      '<dc:title>GSTR-1 ' + xmlEsc(res.period) + '</dc:title><dc:creator>GSTR-1 Converter</dc:creator></cp:coreProperties>');
    files['docProps/app.xml'] = fflate.strToU8('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties"><Application>Microsoft Excel</Application></Properties>');
    files['xl/workbook.xml'] = fflate.strToU8('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">' +
      '<bookViews><workbookView activeTab="0"/></bookViews><sheets>' + sheetsXml + '</sheets><calcPr calcId="191029"/></workbook>');
    files['xl/_rels/workbook.xml.rels'] = fflate.strToU8('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' + relsXml + '</Relationships>');
    files['xl/styles.xml'] = fflate.strToU8(STYLES_XML);
    files['xl/sharedStrings.xml'] = fflate.strToU8(sst.xml());
    return fflate.zipSync(files, { level: 6 });
  }

  /* Validation report as CSV text. */
  function issuesCsv(res) {
    var lines = [['Level', 'Check', 'Count', 'Examples'].join(',')];
    res.issues.forEach(function (i) {
      lines.push([i.level, '"' + i.msg.replace(/"/g, '""') + '"', i.count, '"' + i.refs.join('; ').replace(/"/g, '""') + '"'].join(','));
    });
    return lines.join('\r\n');
  }

  /* Read a GST portal error report (returns_..._errorReport.zip or its .json) and pull out the
     corrections the converter can apply. */
  function parseErrorReport(u8, fileName) {
    var text;
    if (u8[0] === 0x50 && u8[1] === 0x4B) {
      var files = fflate.unzipSync(u8);
      var name = Object.keys(files).filter(function (n) { return /\.json$/i.test(n); })[0];
      if (!name) throw new Error('No .json file inside the error report zip.');
      text = decode(files[name]);
    } else text = decode(u8);
    var j = JSON.parse(text);
    var er = j.error_report || j;
    var out = { gstin: j.gstin || '', fp: j.fp || '', cancelled: [], invalidHsn: [], zeroDocs: [], other: [], total: 0 };
    var seen = {};
    function push(arr, v) { if (v && !seen[arr + v]) { seen[arr + v] = 1; out[arr].push(v); } }
    Object.keys(er).forEach(function (sec) {
      (er[sec] || []).forEach(function (e) {
        var cd = e.error_cd || '', msg = e.error_msg || '';
        if (!cd && !msg) return;
        out.total++;
        if (cd === 'RET191360' || /gstin.*cancel/i.test(msg)) push('cancelled', e.ctin);
        else if (cd === 'RET191349' || /hsn.*not valid/i.test(msg)) {
          ['hsn_b2b', 'hsn_b2c', 'data'].forEach(function (t) { (e[t] || []).forEach(function (h) { push('invalidHsn', h.hsn_sc); }); });
        } else if (cd === 'RET191239' || /zero value/i.test(msg)) {
          (e.nt || e.inv || []).forEach(function (n) { push('zeroDocs', n.nt_num || n.inum); });
        } else {
          var docsN = (e.nt || e.inv || []).map(function (n) { return n.nt_num || n.inum; }).filter(Boolean);
          out.other.push({ section: sec, code: cd, msg: msg, refs: docsN.concat(e.ctin ? [e.ctin] : []) });
        }
      });
    });
    return out;
  }

  function guessStateFromName(name) {
    var toks = String(name || '').toUpperCase().split(/[^A-Z0-9]+/);
    for (var i = toks.length - 1; i >= 0; i--) {
      var t = toks[i];
      if (t.length === 2 && /^[A-Z]{2}$/.test(t)) {
        for (var k = 0; k < STATES.length; k++) if (STATES[k][2] === t && +STATES[k][0] < 90) return STATES[k][0];
      }
    }
    return '';
  }

  var api = {
    loadSource: loadSource, build: build, buildJson: buildJson, splitJson: splitJson, buildExcel: buildExcel,
    issuesCsv: issuesCsv, defaultOptions: defaultOptions, STATES: STATES, RATES: RATES, posLabel: posLabel,
    gstinChecksumOk: gstinChecksumOk, guessStateFromName: guessStateFromName, parseErrorReport: parseErrorReport, _parseDate: parseDate, _openXlsx: openXlsx
  };
  if (typeof module !== 'undefined' && module.exports) module.exports = api;
  else root.GSTR1 = api;
})(typeof self !== 'undefined' ? self : this);
