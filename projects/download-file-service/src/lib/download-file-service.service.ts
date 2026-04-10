import { Injectable } from '@angular/core';
import * as XLSX from 'xlsx';
import jsPDF from 'jspdf';
import autoTable, { RowInput } from 'jspdf-autotable';
declare var google: any;

// ✅ Dependency Interface
export interface DownloadDependencies {
  utilService: any;
  messageService: any;
  documentScanService: any;
  appString: any;
}

export class DownloadFileService {

  private deps!: DownloadDependencies;

  constructor() { }

  setDependencies(deps: DownloadDependencies) {
    this.deps = deps;
  }

  async getAccessToken(): Promise<{ accessToken: string, email: string }> {
    return new Promise((resolve, reject) => {
      const tokenClient = google.accounts.oauth2.initTokenClient({
        client_id: '762787167365-ctsihlroe2f08qgu82ot0vhbp291qn7l.apps.googleusercontent.com',
        scope: 'https://www.googleapis.com/auth/drive.file email profile',
        callback: (resp: any) => {
          if (resp.access_token) {
            // Fetch user info with access token
            fetch('https://www.googleapis.com/oauth2/v3/userinfo', {
              headers: { Authorization: `Bearer ${resp.access_token}` }
            })
              .then(res => res.json())
              .then(profile => {
                localStorage.setItem('google_email', profile.email);
                localStorage.setItem('googleAccessToken', resp.access_token);
                resolve({ accessToken: resp.access_token, email: profile.email });
              })
              .catch(err => reject(err));
          } else {
            reject('No access token');
          }
        },
      });

      tokenClient.requestAccessToken();
    });
  }

  // common csv download formate
  async downloadCSVFormate(doc_id: string, csvContent: any, partyName: string, invoiceNumber: string) {
    const driveEmail = localStorage.getItem('google_email');
    const now = new Date();
    const hh = String(now.getHours()).padStart(2, "0");
    const mm = String(now.getMinutes()).padStart(2, "0");
    const ss = String(now.getSeconds()).padStart(2, "0");
    const fileName = `${partyName}-${invoiceNumber}-${hh}-${mm}-${ss}.csv`;
    const blob = new Blob([csvContent], { type: 'text/csv;charset=utf-8;' });
    this.deps.documentScanService.trackDownload(doc_id, 'csv').subscribe((data: any) => {
      this.deps.documentScanService.updateIsDownload(doc_id).subscribe(async (data: any) => {
        if (!driveEmail) {
          // Local download
          const url = URL.createObjectURL(blob);
          const a = document.createElement('a');
          a.href = url;
          a.download = fileName;
          a.click();
          URL.revokeObjectURL(url);
          return;
        }

        // Google Drive upload
        await this.uploadToDrive(blob, fileName, "text/csv");
      });
    });
  }

  // download csv
  downloadCSV(doc_id: string, tableData: any[] = [], partyName: string, invoiceNumber: string): void {
    const csvHeaders = Object.keys(tableData[0]).join(',') + '\n';
    const csvRows = tableData
      .map((row) => Object.values(row).join(','))
      .join('\n');
    const csvContent = csvHeaders + csvRows;
    this.downloadCSVFormate(doc_id, csvContent, partyName, invoiceNumber);

  }

  // download excel

  async downloadExcel(doc_id: string, tableData: any[] = [], partyName: string, invoiceNumber: string): Promise<void> {
    const driveEmail = localStorage.getItem('google_email');
    const ws: XLSX.WorkSheet = XLSX.utils.json_to_sheet(tableData);
    const wb: XLSX.WorkBook = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, 'Sheet1');
    const now = new Date();
    const hh = String(now.getHours()).padStart(2, "0");
    const mm = String(now.getMinutes()).padStart(2, "0");
    const ss = String(now.getSeconds()).padStart(2, "0");
    const fileName = `${partyName}-${invoiceNumber}-${hh}-${mm}-${ss}.xlsx`;
    this.deps.documentScanService.trackDownload(doc_id, 'xlsx').subscribe((data: any) => {

      this.deps.documentScanService.updateIsDownload(doc_id).subscribe(async (data: any) => {
        if (!driveEmail) {
          XLSX.writeFile(wb, fileName);
        }
        const wbout = XLSX.write(wb, { bookType: 'xlsx', type: 'array' });
        const excelBlob = new Blob([wbout], {
          type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
        });
        await this.uploadToDrive(excelBlob, fileName, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
      });
    });
  }


  // download pdf
  async downloadPDF(doc_id: string, tableData: any[] = [], partyName: string, invoiceNumber: string) {
    const now = new Date();
    const hh = String(now.getHours()).padStart(2, "0");
    const mm = String(now.getMinutes()).padStart(2, "0");
    const ss = String(now.getSeconds()).padStart(2, "0");
    const driveEmail = localStorage.getItem('google_email');

    const doc = new jsPDF();

    const headers = [Object.keys(tableData[0])];

    const rows: RowInput[] = tableData.map((obj) => {
      const rowValues = Object.values(obj).map((val) => {
        if (val instanceof Date) {
          return val.toLocaleDateString();
        }
        return val !== null && val !== undefined ? val.toString() : '';
      });
      return rowValues as RowInput;
    });

    autoTable(doc, {
      head: headers,
      body: rows,
      startY: 10,
      styles: {
        lineWidth: 0.2,
        lineColor: [0, 0, 0],
      },
      headStyles: {
        fillColor: [230, 230, 230],
        textColor: 20,
        halign: 'center',
        valign: 'middle',
      },
      bodyStyles: {
        halign: 'left',
        valign: 'middle',
      },
      theme: 'grid',
    });


    const fileName = `${partyName}-${invoiceNumber}-${hh}-${mm}-${ss}.pdf`;
    this.deps.documentScanService.trackDownload(doc_id, 'pdf').subscribe((data: any) => {

      this.deps.documentScanService.updateIsDownload(doc_id).subscribe(async (data: any) => {
        if (!driveEmail) {
          doc.save(fileName);
          return;
        }
        const pdfBlob = doc.output('blob');
        await this.uploadToDrive(pdfBlob, fileName, "application/pdf");
      });
    });
  }

  //------------------------------ custom csv -----------------------------

  // Normalize each row based on keyMapper.
  // # keyMapper stands for HSN : [ HSN ,HSN Code ]
  // #originalRow stands for gemini response
  normalizeRows(inputRows: any[], keyMapper: any): any[] {
    return inputRows.map((originalRow) => {
      const normalized: any = {};
      for (const [key, value] of Object.entries(originalRow)) {
        const originalKey = key.trim();
        let normalizedKey = originalKey;
        for (const [normKey, possibleKeys] of Object.entries(keyMapper)) {
          if (
            Array.isArray(possibleKeys) &&
            possibleKeys.includes(originalKey)
          ) {
            normalizedKey = normKey;
            break;
          }
        }
        normalized[normalizedKey] = value;
      }
      return normalized;
    });
  }

  downloadTechnomax(ptr: any) {
    return parseFloat(
      this.getValue(ptr?.toString().replace(/,/g, '.'))
    );
  }

  downloadAIOCD(pts: any) {
    return parseFloat(
      this.getValue(pts?.toString().replace(/,/g, '.'))
    );
  }

  // custom csv template
  downloadCustomTemplate(doc_id: string, tableData: any[] = [], columnMapping: any, partyDetails: any, additionalInfo: any, template: any, org_company: string) {
    const orgAccId = localStorage.getItem('org_acc_id')
    let csvContent = '';
    const normalizedRows: any[] = this.normalizeRows(
      tableData,
      columnMapping
    );

    // header data
    const headerData = Object.entries(template['header']).map(
      ([key, value]) => {
        switch (value) {
          case 'H':
            return 'H';
          case '1':
            return "1";
          case '0':
            return '0';
          case 'INVOICE NO':
            return partyDetails['invoice_number'] || '';
          case 'INVOICE DATE':
            return this.formatInvoiceDate(partyDetails['invoice_date']);
          default:
            return '';
        }
      }
    );
    csvContent += headerData.join(',') + '\n';

    // row data
    for (const row of normalizedRows) {
      let taxableAmt = 0;
      let cgst = 0.0;
      let sgst = 0.0;
      const gst = this.getValue(
        row['GST']?.toString().replace(/,/g, '.') ?? ''
      );
      console.log(gst);
      const gstvalue = (parseFloat(gst) || 0) / 2;
      console.log(gstvalue);
      const cgstAmt = row['CGST AMT'];
      const sgstAmt = row['SGST AMT'];

      let ptr = 0.0;
      if (orgAccId === this.deps.appString.AIOCD || orgAccId === this.deps.appString.TECHNOMAX_STANDARD) {
        ptr = this.downloadAIOCD(row['PTS'] === undefined ? row['PTR'] : row['PTS'])
      } else {
        ptr = this.downloadTechnomax(row['PTR'])
      }
      const schAmt = parseFloat(
        this.getValue(row['SCH AMT']?.toString().replace(/,/g, '.'))
      );
      const qty = parseFloat(
        this.splitDigitsByPlus(row['QTY']?.toString().replace(/,/g, '.'))[
        'before_input'
        ]
      );
      let netRate =
        this.getValue(row['SCH AMT']) !== '0' ? ptr - schAmt / qty : ptr;

      let disPer: number = 0;
      const discount = parseFloat(
        this.splitDigitsByPlus(row['DISCOUNT']?.toString().replace(/,/g, '.').replace(this.deps.utilService.NUMBER_REGEX, ''))[
        'after_input'
        ]
      );
      const discountAmt = parseFloat(
        this.getValue(row['DIS AMT']?.toString().replace(/,/g, '.'))
      );
      const discountInPer = row['DISCOUNT']?.toString().replace(/,/g, '.').replace(this.deps.utilService.NUMBER_REGEX, '');
      const discountInAmt = row['DIS AMT']?.toString().replace(/,/g, '.');

      if (!discountInPer || discountInPer === this.deps.appString.NULL) {
        disPer =
          !discountInAmt || discountInAmt === this.deps.appString.NULL
            ? 0.0
            : (discountAmt / (ptr * qty)) * 100;
      } else {
        disPer = discount;
      }

      const tableData = Object.entries(template['body']).map(
        ([_, value]) => {
          switch (value) {
            case 'T':
            case 'H':
            case '0':
              return value;
            case 'PARTY NAME':
              return this.getValue(partyDetails['party_name']?.toString().replace(/,/g, '') ?? '0');
            case 'ITEM CODE':
              return this.getValue(row['ITEM CODE']);
            case 'ITEM COMPANY INFO':
              return this.getValue(row['ITEM COMPANY INFO']);
            case 'ITEM DESCRIPTION':
              return this.getValue(
                row['ITEM DESCRIPTION']?.toString().replace(/,/g, ' ')
              );
            case 'PACK':
              return this.getValue(row['PACK']);
            case 'COMPANY NAME':
              return this.getValue(row['COMPANY NAME']);
            case 'TAXABLE AMT':
              taxableAmt = netRate * qty * (1 - disPer / 100);
              return taxableAmt.toFixed(2);
            case 'BATCH':
              return this.getValue(row['BATCH']);
            case 'PRODUCT CODE':
              return this.getValue(row['PRODUCT CODE']);
            case 'EXP DATE':
              return this.formatExpiryDate(row['EXP DATE']);
            case 'PTR':
            case 'NET RATE':
              let ptrValue = 0;
              if (orgAccId === this.deps.appString.AIOCD || orgAccId === this.deps.appString.TECHNOMAX_STANDARD) {
                // USED ONLY FOR ORG ACCOUNT CATEGORY "AIOCD" INVOICE
                // Applicable for AIOCD category PRM & COMPANY invoices only
                if (org_company === this.deps.appString.PRM_AND_COMPANY || this.getValue(partyDetails['party_name']).toUpperCase() === this.deps.appString.NAME_OF_PRM_AND_COMPANY) {
                  // ptrValue = pts + ((pts*20)/100)
                  ptrValue = Number(this.getValue(row['PTS'])) + ((Number(this.getValue(row['PTS'])) * 20) / 100);
                } else {
                  ptrValue = Number(this.getValue(row['PTR'])) || 0;
                }
              } else {
                ptrValue = Number(this.getValue(row['PTR'])) || 0;
              }
              return ptrValue;
            case 'PTS':
              let ptsValue = 0;
              if (orgAccId === this.deps.appString.AIOCD || orgAccId === this.deps.appString.TECHNOMAX_STANDARD) {
                // USED ONLY FOR ORG ACCOUNT CATEGORY "AIOCD" INVOICE
                // Applicable for AIOCD category CHARAK PHARMA PVT. LTD. invoices only
                if (org_company === this.deps.appString.CHARAK_PHARMA_PVT_LTD || this.getValue(partyDetails['party_name']).toUpperCase() === this.deps.appString.NAME_OF_CHARAK_PHARMA_PVT_LTD) {
                  // ptsValue = TAXABLE AMT / qty
                  ptsValue = qty > 0 ? (Number(this.getValue(row['TAXABLE AMT'])) / qty) : 0;
                } else {
                  ptsValue = Number(this.getValue(row['PTS'])) || 0;
                }
              } else {
                ptsValue = Number(this.getValue(row['PTS'])) || 0;
              }
              return ptsValue;
            case 'MRP':
              return this.getValue(row['MRP'].replace(this.deps.utilService.NUMBER_REGEX, ''));
            case 'QTY':
              return this.splitDigitsByPlus(
                row['QTY']?.toString().replace(/,/g, '.').replace(this.deps.utilService.NUMBER_REGEX, '')
              )['before_input'];
            case 'DISCOUNT':
              return disPer;
            case 'FREE':
              const qtySplit = this.splitDigitsByPlus(row['QTY']);
              return row['QTY']?.toString().includes('+')
                ? qtySplit['after_input']
                : this.getValue(row['FREE']);
            case 'CGST': {
              const gstPer = this.isInvalidField(row['SGST']) ? gst : row['SGST'];
              const cgstper = row['CGST']?.toString().replace(/,/g, '.');
              console.log(gstPer);
              console.log(cgstper);
              if (this.matchesGstPattern(gstPer, this.deps.utilService.decodeGst9Regex) || gstPer === '18') {
                cgst = 9.0;
              } else if (this.matchesGstPattern(gstPer, this.deps.utilService.decodeGst6Regex) || gstPer === '12') {
                cgst = 6.0;
              } else if (
                this.matchesGstPattern(gstPer, this.deps.utilService.decodeGst5Regex) || gstPer === '5') {
                cgst = 2.5;
              } else {
                cgst =
                  this.isInvalidField(row['SGST'])
                    ? gstvalue
                    : parseFloat(cgstper) || 0.0;
              }
              return cgst.toFixed(2);
            }
            case 'CGST AMT':
              // if (cgstAmt && cgstAmt.toString().trim() !== '') {
              //   return cgstAmt.toString().replace(/,/g, '.');
              // }
              return ((taxableAmt * cgst) / 100).toFixed(2);
            case 'SGST': {
              const gstPer = this.isInvalidField(row['SGST']) ? gst : row['SGST'];
              const sgstPer = row['SGST']?.toString().replace(/,/g, '.');
              if (this.matchesGstPattern(gstPer, this.deps.utilService.decodeGst9Regex) || gstPer === '18') {
                sgst = 9.0;
              } else if (
                this.matchesGstPattern(gstPer, this.deps.utilService.decodeGst6Regex) || gstPer === '12') {
                sgst = 6.0;
              } else if (
                this.matchesGstPattern(gstPer, this.deps.utilService.decodeGst5Regex) || gstPer === '5') {
                sgst = 2.5;
              } else {
                sgst =
                  this.isInvalidField(row['SGST'])
                    ? gstvalue
                    : parseFloat(sgstPer) || 0.0;
              }
              return sgst.toFixed(2);
            }
            case 'SGST AMT':
              // if (sgstAmt && sgstAmt.toString().trim() !== '') {
              //   return sgstAmt.toString().replace(/,/g, '.');
              // }
              return ((taxableAmt * sgst) / 100).toFixed(2);
            case 'HSN':
              return this.getValue(row['HSN']);
            default:
              return '0';
          }
        }
      );

      csvContent += tableData.join(',') + '\n';
    }

    //footer data
    const footerData: string[] = Object.entries(template['footer']).map(([_, value]) => {
      switch (value) {
        case 'F':
          return 'F';
        case '0':
          return '0';
        case 'ITEM AMT':
          return `${additionalInfo['item'] ?? 0}`;
        case 'CGST AMT':
          return (additionalInfo['cgstamount']?.toString().replace(/,/g, '') ?? '0');
        case 'SGST AMT':
          return (additionalInfo['sgstamount']?.toString().replace(/,/g, '') ?? '0');
        case 'ROUND OFF':
          return `${additionalInfo['round_off'] ?? 0}`;
        case 'NET PAYABLE':
          return (additionalInfo['netpayable']?.toString().replace(/,/g, '') ?? '0');
        case 'GRAND TOTAL AMT':
          return (additionalInfo['grandtotalamount']?.toString().replace(/,/g, '') ?? '0');
        default:
          return '';
      }
    });
    csvContent += footerData.join(',') + '\n';

    this.downloadCSVFormate(doc_id, csvContent, partyDetails['party_name'], partyDetails['invoice_number']);
  }



  downloadVrajCustomTemplate(doc_id: string, tableData: any[] = [], columnMapping: any, partyDetails: any, additionalInfo: any, template: any, org_company: string) {
    let csvContent = '';
    const normalizedRows: any[] = this.normalizeRows(
      tableData,
      columnMapping
    );
    const titleData = Object.entries(template['header']).map(
      ([key, value]) => {
        switch (value) {
          case '0':
          default:
            return value;
        }
      }
    );
    csvContent += titleData.join(',') + '\n';
    const firstRowOnly = (value: any, i: number) => {
      return i === 0 ? this.getValue(value) : '';
    };
    // row data
    for (let i = 0; i < normalizedRows.length; i++) {
      const row = normalizedRows[i];
      const tableData = Object.entries(template['footer']).map(
        ([_, value]) => {
          switch (value) {
            case '0':
              return value;
            case 'Supplier':
              return this.getValue(partyDetails['party_name']?.toString().replace(/,/g, ' '));
            case 'INVOICE #':
              return `="${this.getValue(
                partyDetails['invoice_number']?.toString().replace(/[,#]/g, '')
              )}"`
            case 'PO Number':
              return this.getValue(partyDetails['po_number']?.toString().replace(/,/g, ' '));
            case 'Order Date':
              return this.getValue(partyDetails['order_date']?.toString().replace(/,/g, ' '));
            case 'Invoice Date':
              return this.getValue(partyDetails['invoice_date']?.toString().replace(/,/g, ' '));
            case 'Ship Date':
              return this.getValue(partyDetails['ship_date']);
            case 'UPC/NDC':
              return this.getValue(row['UPC/NDC']);
            case 'ITEM DESCRIPTION':
              return this.getValue(
                row['ITEM DESCRIPTION']?.toString().replace(/,/g, ' ')
              );
            case 'QTY':
              return this.getValue(row['QTY']?.toString().replace(this.deps.utilService.NUMBER_REGEX, ''));
            case 'PACK':
              return this.getValue(row['PACK']);
            case 'PTR':
              return this.getValue(row['PTR']?.toString().replace(this.deps.utilService.NUMBER_REGEX, ''));
            case 'Total Cost':
              return this.getValue(row['Total Cost']?.toString().replace(this.deps.utilService.NUMBER_REGEX, ''));
            case 'SHIPPING CHARGES':
              return firstRowOnly(this.getAdditionalInfoValue(additionalInfo, 'shipping_charge'), i);
            case 'FUEL SURCHARGE':
              return firstRowOnly(this.getAdditionalInfoValue(additionalInfo, 'fuel_charge'), i);
            case 'TAX':
              return firstRowOnly(this.getAdditionalInfoValue(additionalInfo, 'tax_amount'), i);
            case 'VAT':
              return firstRowOnly(this.getAdditionalInfoValue(additionalInfo, 'vat_amount'), i);
            case 'ADDITIONAL CHARGES':
              return firstRowOnly(this.getAdditionalInfoValue(additionalInfo, 'additional_charge'), i);
            case 'CREDIT CARD CHARGES':
              return firstRowOnly(this.getAdditionalInfoValue(additionalInfo, 'credit_card_charge'), i);
            case 'TARIFF SURCHARGE':
              return firstRowOnly(this.getAdditionalInfoValue(additionalInfo, 'tariff_charge'), i);
            case 'HANDLING':
              return firstRowOnly(this.getAdditionalInfoValue(additionalInfo, 'handling_charge'), i);
            case 'MISC CHARGES':
              return firstRowOnly(this.getAdditionalInfoValue(additionalInfo, 'misc_charge'), i);
            case 'DEPOSIT':
              return firstRowOnly(this.getAdditionalInfoValue(additionalInfo, 'deposit_applied'), i);
            case 'DEP APPLIED':
              return firstRowOnly(this.getAdditionalInfoValue(additionalInfo, 'deposit_applied'), i);
            case 'CC FEE':
              return this.getValue(row['CC FEE']);
            case 'FREIGHT ALLOWANCE':
              return firstRowOnly(this.getAdditionalInfoValue(additionalInfo, 'freight_allowance'), i);
            case 'DISCOUNT':
              return firstRowOnly(this.getAdditionalInfoValue(additionalInfo, 'line_discount_amount'), i);
            case 'TOTAL AMOUNT':
              return firstRowOnly(this.getAdditionalInfoValue(additionalInfo, 'final_total_amount'), i);
            default:
              return '0';
          }
        }
      );

      csvContent += tableData.join(',') + '\n';
    }
    this.downloadCSVFormate(doc_id, csvContent, partyDetails['party_name'], partyDetails['invoice_number']);
  }

  getAdditionalInfoValue(additionalInfo: any, key: any) {
    return this.getValue(additionalInfo[key]?.toString().replace(this.deps.utilService.NUMBER_REGEX, ''))
  }

  downloadTramaxCustomTemplate(doc_id: string, tableData: any[] = [], columnMapping: any, partyDetails: any, additionalInfo: any, template: any, busy_type: string) {
    const excelData: any[] = [];
    const normalizedRows: any[] = this.normalizeRows(
      tableData,
      columnMapping
    );
    let totalCgstAmt = 0.0;
    let totalLineCgstAmt = 0.0;
    let totalSgstAmt = 0.0;
    let totalLineSgstAmt = 0.0;
    let totalIgstAmt = 0.0;
    let totalLineIgstAmt = 0.0;
    let totalFreightAmt = 0.0;

    const firstRowOnly = (value: any, i: number) => {
      return i === 0 ? this.getValue(value) : '';
    };

    const lastRowOnly = (value: any, i: number, length: number) => {
      return i === length - 1 ? this.getValue(value) : '';
    };

    const hasLineFreight = normalizedRows.some(row => {
      const lineFreight = parseFloat(this.getValue(row['FREIGHT CHARGES']?.toString().replace(/,/g, '')));
      return !isNaN(lineFreight) && lineFreight > 0;
    });
    // row data
    for (let i = 0; i < normalizedRows.length; i++) {
      const row = normalizedRows[i];
      let cgst = 0.0;
      let sgst = 0.0;
      let igst = 0.0;
      let ptr = 0.0;
      let qty = 0.0;
      let amount = 0.0;
      let dis = 0.0;

      const rowGst = Number(
        row['GST']?.toString().replace(/,/g, '.') ?? 0
      );
      const igstPercent = Number(additionalInfo['igst_percent_value'] ?? 0);
      const isIgst = igstPercent !== 0;
      ptr = parseFloat(this.getValue(row['PTR']?.toString().replace(/,/g, '')));
      qty = parseFloat(this.getValue(row['QTY']?.toString().replace(/,/g, '')));
      dis = parseFloat(this.getValue(row['DISCOUNT']?.toString().replace(/,/g, '')));
      let freight_charges = parseFloat(
        this.getValue(additionalInfo['freight_charges']?.toString().replace(/,/g, ''))
      );
      let line_freight_charges = parseFloat(
        this.getValue(row['FREIGHT CHARGES']?.toString().replace(/,/g, ''))
      );

      freight_charges = isNaN(freight_charges) ? 0 : freight_charges;
      line_freight_charges = isNaN(line_freight_charges) ? 0 : line_freight_charges;
      let final_freight = 0;
      if (hasLineFreight) {
        // If ANY row has freight → use only line-wise
        final_freight = line_freight_charges > 0 ? line_freight_charges : 0;
      } else {
        // If NO row has freight → use global freight
        final_freight = freight_charges;
      }
      amount = ptr * qty;

      const gst = rowGst === 0
        ? Number(additionalInfo['igst_percent_value'] ?? 0)
        : rowGst;
      const gstvalue = (parseFloat(gst.toString()) || 0) / 2;

      const excelRow = Object.entries(template['header']).map(
        ([_, value]) => {
          switch (value) {
            case 'SERIES':
              return 'Main';
            case 'DATE':
              return this.getValue(partyDetails['invoice_date']);
            case 'VCH_NO':
              return partyDetails['vch_number'] || '';
            case 'INVOICE_NO':
              return partyDetails['invoice_number'] || '';
            case 'PURCHASE TYPE':
              return busy_type == this.deps.appString.ITEM_WISE ? 'Local-ItemWise' : isIgst ? 'Central-MultiRate' : 'Local-MultiRate'
            case 'PARTY NAME':
              return partyDetails['select_party_name'] && partyDetails['select_party_name'] !== ''
                ? partyDetails['select_party_name']
                : partyDetails['party_name']
            case 'MATERIAL CENTER':
              return 'Main Store';
            case 'ITEM DESCRIPTION':
              return this.getValue(
                row['ITEM DESCRIPTION']?.toString().replace(/,/g, ' ')
              );
            case 'QTY':
              return qty
            case 'PACK':
              return this.getValue(row['PACK']);
            case 'AMOUNT':
              amount = ptr * qty
              return amount;
            case 'PTR':
              const gross_amt = amount - (amount * dis / 100);
              ptr = gross_amt / qty;
              return ptr;
            case 'HSN':
              return this.getValue(row['HSN']);
            case 'FREIGHT_CHARGES':
              return firstRowOnly(this.getValue(final_freight) === '0' ? '' : 'Freight & Forwarding Charges', i);
            case 'FREIGHT_CHARGES_AMOUNT':
              totalFreightAmt += final_freight;
              return hasLineFreight ? '' : firstRowOnly(final_freight, i);
            case 'ROUND_OFF':
              const roundOff = Number(additionalInfo['roundoffamount'] || 0);
              if (roundOff === 0) {
                return '0';
              } else if (roundOff > 0) {
                return firstRowOnly('Rounded Off (+)', i);
              } else {
                return firstRowOnly('Rounded Off (-)', i);;
              }
            case 'ROUND_OFF_AMT':
              const rawValue = additionalInfo['roundoffamount'] || '0';
              const cleanedValue = String(rawValue).replace(this.deps.utilService.NUMBER_REGEX, '');
              return firstRowOnly(Math.abs(Number(cleanedValue || 0)), i);
            case 'TEX_CATEGORY': {
              let texCategory = '';
              const gst = this.getValue(row['GST']);
              if (gst === '18') {
                texCategory = 'GST 18%'
              } else if (gst === '5') {
                texCategory = 'GST 5%'
              } else {
                texCategory = 'GST 40%'
              }
              return busy_type == this.deps.appString.ITEM_WISE ? texCategory : 0;
            }
            case 'CGST_NAME':
              return busy_type == this.deps.appString.ITEM_WISE ? 0 : !isIgst ? firstRowOnly('CGST', i) : 0;
            case 'SGST_NAME':
              return busy_type == this.deps.appString.ITEM_WISE ? 0 : !isIgst ? firstRowOnly('SGST', i) : 0;
            case 'IGST_NAME':
              return busy_type == this.deps.appString.ITEM_WISE ? 0 : isIgst ? firstRowOnly('IGST', i) : 0;
            case 'CGST': {
              let gstPer: any = 0;

              // 1️⃣ priority: row['CGST']
              if (!this.isInvalidField(row['CGST']) && Number(row['CGST']) !== 0) {
                gstPer = row['CGST'];
              }
              // 2️⃣ fallback: gst (row GST / IGST)
              else if (!this.isInvalidField(gst) && Number(gst) !== 0) {
                gstPer = gst;
              }
              // 3️⃣ fallback: additionalInfo['cgst_percent_value']
              else {
                gstPer = Number((additionalInfo['cgst_percent_value'] ?? '0').toString().replace('%', '')) * 2;
              }

              const cgstper = row['CGST']?.toString().replace(/,/g, '.');

              if (
                this.matchesGstPattern(gstPer, this.deps.utilService.decodeGst9Regex) ||
                gstPer == 18
              ) {
                cgst = 9.0;
              } else if (
                this.matchesGstPattern(gstPer, this.deps.utilService.decodeGst6Regex) ||
                gstPer == 12
              ) {
                cgst = 6.0;
              } else if (
                this.matchesGstPattern(gstPer, this.deps.utilService.decodeGst5Regex) ||
                gstPer == 5
              ) {
                cgst = 2.5;
              } else {
                cgst =
                  !this.isInvalidField(cgstper)
                    ? parseFloat(cgstper)
                    : (Number(gstPer) / 2) || 0;
              }

              return busy_type == this.deps.appString.ITEM_WISE ? 0 : !isIgst ? firstRowOnly(cgst.toFixed(2), i) : 0;
            }
            case 'CGST AMT': {
              totalCgstAmt += busy_type == this.deps.appString.ITEM_WISE ? 0 : isIgst ? 0 : amount;
              if (!hasLineFreight && lastRowOnly(final_freight, i, normalizedRows.length)) {
                totalCgstAmt = ((totalCgstAmt + final_freight) * cgst) / 100;
              } else if (hasLineFreight) {
                let tax_amount = amount + final_freight;
                const cgstAmt = (tax_amount * cgst) / 100;
                totalLineCgstAmt += busy_type == this.deps.appString.ITEM_WISE ? 0 : isIgst ? 0 : cgstAmt;
              }
              return '';
            }
            case 'SGST': {
              let gstPer: any = 0;

              // 1️⃣ priority: row['SGST']
              if (!this.isInvalidField(row['SGST']) && Number(row['SGST']) !== 0) {
                gstPer = row['SGST'];
              }
              // 2️⃣ fallback: gst (row GST / IGST)
              else if (!this.isInvalidField(gst) && Number(gst) !== 0) {
                gstPer = gst;
              }
              // 3️⃣ fallback: additionalInfo['sgst_percent_value']
              else {
                gstPer = Number((additionalInfo['sgst_percent_value'] ?? '0').toString().replace('%', '')) * 2;
              }

              const sgstPer = row['SGST']?.toString().replace(/,/g, '.');

              if (
                this.matchesGstPattern(gstPer, this.deps.utilService.decodeGst9Regex) ||
                gstPer == 18
              ) {
                sgst = 9.0;
              } else if (
                this.matchesGstPattern(gstPer, this.deps.utilService.decodeGst6Regex) ||
                gstPer == 12
              ) {
                sgst = 6.0;
              } else if (
                this.matchesGstPattern(gstPer, this.deps.utilService.decodeGst5Regex) ||
                gstPer == 5
              ) {
                sgst = 2.5;
              } else {
                sgst =
                  !this.isInvalidField(sgstPer)
                    ? parseFloat(sgstPer)
                    : (Number(gstPer) / 2) || 0;
              }

              return busy_type == this.deps.appString.ITEM_WISE ? 0 : !isIgst ? firstRowOnly(sgst.toFixed(2), i) : 0;
            }
            case 'IGST':
              igst = Number(additionalInfo['igst_percent_value']);
              return busy_type == this.deps.appString.ITEM_WISE ? 0 : isIgst ? firstRowOnly(igst.toFixed(2), i) : 0;

            case 'SGST AMT': {
              totalSgstAmt += busy_type == this.deps.appString.ITEM_WISE ? 0 : isIgst ? 0 : amount;
              if (!hasLineFreight && lastRowOnly(final_freight, i, normalizedRows.length)) {
                totalSgstAmt = ((totalSgstAmt + final_freight) * sgst) / 100;
              } else if (hasLineFreight) {
                let tax_amount = amount + final_freight;
                const sgstAmt = (tax_amount * sgst) / 100;
                totalLineSgstAmt += busy_type == this.deps.appString.ITEM_WISE ? 0 : isIgst ? 0 : sgstAmt;
              }
              return '';
            }
            case 'IGST AMT': {
              const igst = Number(additionalInfo['igst_percent_value'] ?? 0);
              totalIgstAmt += busy_type == this.deps.appString.ITEM_WISE ? 0 : !isIgst ? 0 : amount;
              if (!hasLineFreight && lastRowOnly(final_freight, i, normalizedRows.length)) {
                totalIgstAmt = ((totalIgstAmt + final_freight) * igst) / 100;
              } else if (hasLineFreight) {
                let tax_amount = amount + final_freight;
                const igstAmt = (tax_amount * igst) / 100;
                totalLineIgstAmt += busy_type == this.deps.appString.ITEM_WISE ? 0 : !isIgst ? 0 : igstAmt;
              }
              return '';
            }
            default:
              return '0';
          }
        }
      );
      excelData.push(excelRow);
    }
    if (excelData.length > 0) {

      const headers = Object.values(template['header']);

      const cgstIndex = headers.indexOf('CGST AMT');
      const sgstIndex = headers.indexOf('SGST AMT');
      const igstIndex = headers.indexOf('IGST AMT');
      const freightChargesIndex = headers.indexOf('FREIGHT_CHARGES_AMOUNT');

      if (cgstIndex !== -1) {
        excelData[0][cgstIndex] = hasLineFreight ? totalLineCgstAmt.toFixed(2) : totalCgstAmt.toFixed(2);
      }

      if (sgstIndex !== -1) {
        excelData[0][sgstIndex] = hasLineFreight ? totalLineSgstAmt.toFixed(2) : totalSgstAmt.toFixed(2);
      }

      if (igstIndex !== -1) {
        excelData[0][igstIndex] = hasLineFreight ? totalLineIgstAmt.toFixed(2) : totalIgstAmt.toFixed(2);
      }

      if (hasLineFreight && freightChargesIndex !== -1) {
        excelData[0][freightChargesIndex] = totalFreightAmt.toFixed(2);
      }
    }
    this.downloadExcelFile(doc_id, excelData, partyDetails['party_name'], partyDetails['invoice_number']);
  }

  downloadBulkBusyCustomTemplate(key_value_json: any, busy_type: string, columnMapping: any, template: any) {
    const excelData: any[] = [];
    for (const doc of key_value_json) {
      const tableData = doc.key_value_json.table_data;
      const partyDetails = doc.key_value_json.text_data;
      const additionalInfo = doc.key_value_json.additional_info;
      const normalizedRows: any[] = this.normalizeRows(
        tableData,
        columnMapping
      );
      let totalCgstAmt = 0.0;
      let totalLineCgstAmt = 0.0;
      let totalSgstAmt = 0.0;
      let totalLineSgstAmt = 0.0;
      let totalIgstAmt = 0.0;
      let totalLineIgstAmt = 0.0;
      let totalFreightAmt = 0.0;

      const firstRowOnly = (value: any, i: number) => {
        return i === 0 ? this.getValue(value) : '';
      };

      const lastRowOnly = (value: any, i: number, length: number) => {
        return i === length - 1 ? this.getValue(value) : '';
      };

      const hasLineFreight = normalizedRows.some(row => {
        const lineFreight = parseFloat(this.getValue(row['FREIGHT CHARGES']?.toString().replace(/,/g, '')));
        return !isNaN(lineFreight) && lineFreight > 0;
      });
      // row data
      for (let i = 0; i < normalizedRows.length; i++) {
        const row = normalizedRows[i];
        let cgst = 0.0;
        let sgst = 0.0;
        let igst = 0.0;
        let ptr = 0.0;
        let qty = 0.0;
        let amount = 0.0;
        let dis = 0.0;

        const rowGst = Number(
          row['GST']?.toString().replace(/,/g, '.') ?? 0
        );
        const igstPercent = Number(additionalInfo['igst_percent_value'] ?? 0);
        const isIgst = igstPercent !== 0;
        ptr = parseFloat(this.getValue(row['PTR']?.toString().replace(/,/g, '')));
        qty = parseFloat(this.getValue(row['QTY']?.toString().replace(/,/g, '')));
        dis = parseFloat(this.getValue(row['DISCOUNT']?.toString().replace(/,/g, '')));
        let freight_charges = parseFloat(
          this.getValue(additionalInfo['freight_charges']?.toString().replace(/,/g, ''))
        );
        let line_freight_charges = parseFloat(
          this.getValue(row['FREIGHT CHARGES']?.toString().replace(/,/g, ''))
        );

        freight_charges = isNaN(freight_charges) ? 0 : freight_charges;
        line_freight_charges = isNaN(line_freight_charges) ? 0 : line_freight_charges;
        let final_freight = 0;
        if (hasLineFreight) {
          // If ANY row has freight → use only line-wise
          final_freight = line_freight_charges > 0 ? line_freight_charges : 0;
        } else {
          // If NO row has freight → use global freight
          final_freight = freight_charges;
        }
        amount = ptr * qty;

        const gst = rowGst === 0
          ? Number(additionalInfo['igst_percent_value'] ?? 0)
          : rowGst;
        const gstvalue = (parseFloat(gst.toString()) || 0) / 2;

        const excelRow = Object.entries(template['header']).map(
          ([_, value]) => {
            switch (value) {
              case 'SERIES':
                return 'Main';
              case 'DATE':
                return this.getValue(partyDetails['invoice_date']);
              case 'VCH_NO':
                return partyDetails['vch_number'] || '';
              case 'INVOICE_NO':
                return partyDetails['invoice_number'] || '';
              case 'PURCHASE TYPE':
                return busy_type == this.deps.appString.ITEM_WISE ? 'Local-ItemWise' : isIgst ? 'Central-MultiRate' : 'Local-MultiRate'
              case 'PARTY NAME':
                return partyDetails['select_party_name'] && partyDetails['select_party_name'] !== ''
                  ? partyDetails['select_party_name']
                  : partyDetails['party_name']
              case 'MATERIAL CENTER':
                return 'Main Store';
              case 'ITEM DESCRIPTION':
                return this.getValue(
                  row['ITEM DESCRIPTION']?.toString().replace(/,/g, ' ')
                );
              case 'QTY':
                return qty
              case 'PACK':
                return this.getValue(row['PACK']);
              case 'AMOUNT':
                amount = ptr * qty
                return amount;
              case 'PTR':
                const gross_amt = amount - (amount * dis / 100);
                ptr = gross_amt / qty;
                return ptr;
              case 'HSN':
                return this.getValue(row['HSN']);
              case 'FREIGHT_CHARGES':
                return firstRowOnly(this.getValue(final_freight) === '0' ? '' : 'Freight & Forwarding Charges', i);
              case 'FREIGHT_CHARGES_AMOUNT':
                totalFreightAmt += final_freight;
                return hasLineFreight ? '' : firstRowOnly(final_freight, i);
              case 'ROUND_OFF':
                const roundOff = Number(additionalInfo['roundoffamount'] || 0);
                if (roundOff === 0) {
                  return '0';
                } else if (roundOff > 0) {
                  return firstRowOnly('Rounded Off (+)', i);
                } else {
                  return firstRowOnly('Rounded Off (-)', i);;
                }
              case 'ROUND_OFF_AMT':
                const rawValue = additionalInfo['roundoffamount'] || '0';
                const cleanedValue = String(rawValue).replace(this.deps.utilService.NUMBER_REGEX, '');
                return firstRowOnly(Math.abs(Number(cleanedValue || 0)), i);
              case 'TEX_CATEGORY': {
                let texCategory = '';
                const gst = this.getValue(row['GST']);
                if (gst === '18') {
                  texCategory = 'GST 18%'
                } else if (gst === '5') {
                  texCategory = 'GST 5%'
                } else {
                  texCategory = 'GST 40%'
                }
                return busy_type == this.deps.appString.ITEM_WISE ? texCategory : 0;
              }
              case 'CGST_NAME':
                return busy_type == this.deps.appString.ITEM_WISE ? 0 : !isIgst ? firstRowOnly('CGST', i) : 0;
              case 'SGST_NAME':
                return busy_type == this.deps.appString.ITEM_WISE ? 0 : !isIgst ? firstRowOnly('SGST', i) : 0;
              case 'IGST_NAME':
                return busy_type == this.deps.appString.ITEM_WISE ? 0 : isIgst ? firstRowOnly('IGST', i) : 0;
              case 'CGST': {
                let gstPer: any = 0;

                // 1️⃣ priority: row['CGST']
                if (!this.isInvalidField(row['CGST']) && Number(row['CGST']) !== 0) {
                  gstPer = row['CGST'];
                }
                // 2️⃣ fallback: gst (row GST / IGST)
                else if (!this.isInvalidField(gst) && Number(gst) !== 0) {
                  gstPer = gst;
                }
                // 3️⃣ fallback: additionalInfo['cgst_percent_value']
                else {
                  gstPer = Number((additionalInfo['cgst_percent_value'] ?? '0').toString().replace('%', '')) * 2;
                }

                const cgstper = row['CGST']?.toString().replace(/,/g, '.');

                if (
                  this.matchesGstPattern(gstPer, this.deps.utilService.decodeGst9Regex) ||
                  gstPer == 18
                ) {
                  cgst = 9.0;
                } else if (
                  this.matchesGstPattern(gstPer, this.deps.utilService.decodeGst6Regex) ||
                  gstPer == 12
                ) {
                  cgst = 6.0;
                } else if (
                  this.matchesGstPattern(gstPer, this.deps.utilService.decodeGst5Regex) ||
                  gstPer == 5
                ) {
                  cgst = 2.5;
                } else {
                  cgst =
                    !this.isInvalidField(cgstper)
                      ? parseFloat(cgstper)
                      : (Number(gstPer) / 2) || 0;
                }

                return busy_type == this.deps.appString.ITEM_WISE ? 0 : !isIgst ? firstRowOnly(cgst.toFixed(2), i) : 0;
              }
              case 'CGST AMT': {
                totalCgstAmt += busy_type == this.deps.appString.ITEM_WISE ? 0 : isIgst ? 0 : amount;
                if (!hasLineFreight && lastRowOnly(final_freight, i, normalizedRows.length)) {
                  totalCgstAmt = ((totalCgstAmt + final_freight) * cgst) / 100;
                } else if (hasLineFreight) {
                  let tax_amount = amount + final_freight;
                  const cgstAmt = (tax_amount * cgst) / 100;
                  totalLineCgstAmt += busy_type == this.deps.appString.ITEM_WISE ? 0 : isIgst ? 0 : cgstAmt;
                }
                return '';
              }
              case 'SGST': {
                let gstPer: any = 0;

                // 1️⃣ priority: row['SGST']
                if (!this.isInvalidField(row['SGST']) && Number(row['SGST']) !== 0) {
                  gstPer = row['SGST'];
                }
                // 2️⃣ fallback: gst (row GST / IGST)
                else if (!this.isInvalidField(gst) && Number(gst) !== 0) {
                  gstPer = gst;
                }
                // 3️⃣ fallback: additionalInfo['sgst_percent_value']
                else {
                  gstPer = Number((additionalInfo['sgst_percent_value'] ?? '0').toString().replace('%', '')) * 2;
                }

                const sgstPer = row['SGST']?.toString().replace(/,/g, '.');

                if (
                  this.matchesGstPattern(gstPer, this.deps.utilService.decodeGst9Regex) ||
                  gstPer == 18
                ) {
                  sgst = 9.0;
                } else if (
                  this.matchesGstPattern(gstPer, this.deps.utilService.decodeGst6Regex) ||
                  gstPer == 12
                ) {
                  sgst = 6.0;
                } else if (
                  this.matchesGstPattern(gstPer, this.deps.utilService.decodeGst5Regex) ||
                  gstPer == 5
                ) {
                  sgst = 2.5;
                } else {
                  sgst =
                    !this.isInvalidField(sgstPer)
                      ? parseFloat(sgstPer)
                      : (Number(gstPer) / 2) || 0;
                }

                return busy_type == this.deps.appString.ITEM_WISE ? 0 : !isIgst ? firstRowOnly(sgst.toFixed(2), i) : 0;
              }
              case 'IGST':
                igst = Number(additionalInfo['igst_percent_value']);
                return busy_type == this.deps.appString.ITEM_WISE ? 0 : isIgst ? firstRowOnly(igst.toFixed(2), i) : 0;

              case 'SGST AMT': {
                totalSgstAmt += busy_type == this.deps.appString.ITEM_WISE ? 0 : isIgst ? 0 : amount;
                if (!hasLineFreight && lastRowOnly(final_freight, i, normalizedRows.length)) {
                  totalSgstAmt = ((totalSgstAmt + final_freight) * sgst) / 100;
                } else if (hasLineFreight) {
                  let tax_amount = amount + final_freight;
                  const sgstAmt = (tax_amount * sgst) / 100;
                  totalLineSgstAmt += busy_type == this.deps.appString.ITEM_WISE ? 0 : isIgst ? 0 : sgstAmt;
                }
                return '';
              }
              case 'IGST AMT': {
                const igst = Number(additionalInfo['igst_percent_value'] ?? 0);
                totalIgstAmt += busy_type == this.deps.appString.ITEM_WISE ? 0 : !isIgst ? 0 : amount;
                if (!hasLineFreight && lastRowOnly(final_freight, i, normalizedRows.length)) {
                  totalIgstAmt = ((totalIgstAmt + final_freight) * igst) / 100;
                } else if (hasLineFreight) {
                  let tax_amount = amount + final_freight;
                  const igstAmt = (tax_amount * igst) / 100;
                  totalLineIgstAmt += busy_type == this.deps.appString.ITEM_WISE ? 0 : !isIgst ? 0 : igstAmt;
                }
                return '';
              }
              default:
                return '0';
            }
          }
        );
        excelData.push(excelRow);
      }
      if (excelData.length > 0) {
        const startIndex = excelData.length - normalizedRows.length;

        const headers = Object.values(template['header']);

        const cgstIndex = headers.indexOf('CGST AMT');
        const sgstIndex = headers.indexOf('SGST AMT');
        const igstIndex = headers.indexOf('IGST AMT');
        const freightIndex = headers.indexOf('FREIGHT_CHARGES_AMOUNT');

        if (cgstIndex !== -1) {
          excelData[startIndex][cgstIndex] = hasLineFreight ? totalLineSgstAmt.toFixed(2) : totalSgstAmt.toFixed(2);
        }

        if (sgstIndex !== -1) {
          excelData[startIndex][sgstIndex] = hasLineFreight ? totalLineSgstAmt.toFixed(2) : totalSgstAmt.toFixed(2);
        }

        if (igstIndex !== -1) {
          excelData[startIndex][igstIndex] = hasLineFreight ? totalLineIgstAmt.toFixed(2) : totalIgstAmt.toFixed(2);
        }

        if (hasLineFreight && freightIndex !== -1) {
          excelData[startIndex][freightIndex] = totalFreightAmt.toFixed(2);
        }
      }
    }

    this.downloadExcelFile(key_value_json[0]['document_master_id'], excelData, busy_type, "");
  }


  downloadGofrugalTemplate(doc_id: string, tableData: any[] = [], columnMapping: any, partyDetails: any, additionalInfo: any, template: any, org_company: string) {
    let csvContent = '';
    const normalizedRows: any[] = this.normalizeRows(
      tableData,
      columnMapping
    );

    // header data
    const headerData = Object.entries(template['header']).map(
      ([key, value]) => {
        switch (value) {
          case '0':
          default:
            return value;
        }
      }
    );
    csvContent += headerData.join(',') + '\n';

    // row data
    for (const row of normalizedRows) {
      const gst = this.getValue(
        row['GST']?.toString().replace(/,/g, '.') ?? ''
      );

      const tableData = Object.entries(template['body']).map(
        ([_, value]) => {
          switch (value) {
            case 'T':
            case 'H':
            case '0':
              return value;
            case 'ITEM DESCRIPTION':
              return this.getValue(
                row['ITEM DESCRIPTION']?.toString().replace(/,/g, ' ')
              );
            case 'PRODUCT TYPE':
              return this.getValue(
                row['PRODUCT TYPE']?.toString().replace(/,/g, ' ')
              );
            case 'ALIAS':
              return this.getValue(
                row['ALIAS']?.toString().replace(/,/g, ' ')
              );

            case 'PTR':
              return Number(this.getValue(row['PTR'])) || 0;
            case 'PTS':
              return Number(this.getValue(row['PTS'])) || 0;;
            case 'MRP':
              return this.getValue(row['MRP']);
            case 'QTY':
              return this.splitDigitsByPlus(
                row['QTY']?.toString().replace(/,/g, '.')
              )['before_input'];
            case 'NO OF SERIAL FIELDS':
              return this.getValue(
                row['NO OF SERIAL FIELDS']?.toString().replace(/,/g, ' ')
              );
            case 'INCLUSIVE':
              return this.getValue(
                row['INCLUSIVE']?.toString().replace(/,/g, ' ')
              );
            case 'GST': {
              const cgstper = row['CGST']?.toString().replace(/,/g, '.');
              const sgstper = row['SGST']?.toString().replace(/,/g, '.');
              let gstPer = '';
              if (gst !== '0') {
                gstPer = gst;
              } else if (parseFloat(cgstper) > 0) {
                gstPer = this.getValue(parseFloat(cgstper) * 2)
              } else {
                gstPer = this.getValue(parseFloat(sgstper) * 2)
              }
              return gstPer;
            }
            default:
              return '0';
          }
        }
      );

      csvContent += tableData.join(',') + '\n';
    }
    this.downloadCSVFormate(doc_id, csvContent, partyDetails['party_name'], partyDetails['invoice_number']);
  }


  async downloadExcelFile(doc_id: string, tableData: any[] = [], partyName: string, invoiceNumber: string): Promise<void> {
    const driveEmail = localStorage.getItem('google_email');
    const ws: XLSX.WorkSheet = XLSX.utils.aoa_to_sheet(tableData);
    const wb: XLSX.WorkBook = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, 'Sheet1');
    const now = new Date();
    const hh = String(now.getHours()).padStart(2, "0");
    const mm = String(now.getMinutes()).padStart(2, "0");
    const ss = String(now.getSeconds()).padStart(2, "0");
    const fileName = `${partyName}-${invoiceNumber}-${hh}-${mm}-${ss}.xlsx`;
    this.deps.documentScanService.trackDownload(doc_id, 'xlsx').subscribe((data: any) => {

      this.deps.documentScanService.updateIsDownload(doc_id).subscribe(async (data: any) => {
        if (!driveEmail) {
          XLSX.writeFile(wb, fileName);
        }
        const wbout = XLSX.write(wb, { bookType: 'xlsx', type: 'array' });
        const excelBlob = new Blob([wbout], {
          type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
        });
        await this.uploadToDrive(excelBlob, fileName, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
      });
    });
  }

  // invoice date formatting
  formatInvoiceDate(date: string | null | undefined): string {
    if (!date || date.trim() === '') return '';

    const cleaned = date.replace(/[-\/]/g, '');

    if (cleaned.length === 6) {
      const part1 = cleaned.substring(0, 4);
      const year = cleaned.substring(4);
      return `${part1}20${year}`;
    }

    return cleaned;
  }

  // Helper to safely get string values with default '0'
  getValue(value: any): string {
    if (
      value === null ||
      value === undefined ||
      value.toString().toLowerCase() === this.deps.appString.NULL
    ) {
      return '0';
    }

    const str = value.toString().trim();
    return str === '' ? '0' : str;
  }

  // Returns the part before '+' from the given value, trimmed of spaces.
  splitDigitsByPlus(value: any): { before_input: string; after_input: string } {
    if (value === null || value === undefined) {
      return { before_input: '0', after_input: '0' };
    }

    const str = value.toString().trim();

    if (str.includes('+')) {
      const parts = str.split('+');
      return {
        before_input: parts[0].trim(),
        after_input: parts[1].trim(),
      };
    }

    const safeValue = str === '' || str.toLowerCase() === this.deps.appString.NULL ? '0' : str;

    return {
      before_input: safeValue,
      after_input: safeValue,
    };
  }

  // Formatting expiry date to ddMMyyyy format
  // Handle both '-' and '/' delimiters
  // Get last day of the month
  formatExpiryDate(rawValue: string | null | undefined): string {
    if (!rawValue || rawValue.trim() === '') return '';

    try {
      const str = rawValue.trim();
      const separator = str.includes('-')
        ? '-'
        : str.includes('/')
          ? '/'
          : str.includes('.')
            ? '.'
            : str.includes(' ')
              ? ' '
              : null;
      if (!separator) return '';

      const parts = str.split(separator).filter(p => p.trim() !== '');

      let day: number;
      let month: number;
      let year: number;

      // ✅ Case 1: DD-MM-YYYY
      // ✅ Case 1: 3 parts (DD-MM-YYYY OR YYYY-MM-DD)
      if (parts.length === 3) {
        let p1 = parseInt(parts[0]);
        let p2 = parseInt(parts[1]);
        let p3 = parseInt(parts[2]);

        if (isNaN(p1) || isNaN(p2) || isNaN(p3)) return '';

        // 👉 Detect format
        if (parts[0].length === 4) {
          // YYYY-MM-DD
          year = p1;
          month = p2;
          day = p3;
        } else {
          // DD-MM-YYYY
          day = p1;
          month = p2;
          year = p3;
        }

        return (
          day.toString().padStart(2, '0') +
          month.toString().padStart(2, '0') +
          year.toString()
        );
      }

      // ✅ Case 2: MM-YY or MMM-YY
      if (parts.length === 2) {
        let [monthStr, yearStr] = parts;

        month = parseInt(monthStr);
        if (isNaN(month)) {
          const date = new Date(`${monthStr} 1, 2000`);
          if (isNaN(date.getTime())) return '';
          month = date.getMonth() + 1;
        }

        // If year is 2 digits (e.g. "26"), assume it belongs to 2000s (→ 2026)
        // If year is already 4 digits (e.g. "2026"), use it as-is
        year =
          yearStr.length === 2
            ? 2000 + parseInt(yearStr)
            : parseInt(yearStr);

        // Last day of the month
        const lastDay = new Date(year, month, 0);// 0 = last day of previous month

        const dd = lastDay.getDate().toString().padStart(2, '0');
        const mm = (lastDay.getMonth() + 1).toString().padStart(2, '0');
        const yyyy = lastDay.getFullYear().toString();

        return `${dd}${mm}${yyyy}`;
      }

      return '';
    } catch {
      return '';
    }
  }

  // gst percentage Regex check
  matchesGstPattern(input: string, pattern: RegExp): boolean {
    return pattern.test(input);
  }

  // Checks if a given value is considered invalid or empty.
  isInvalidField(val: any): boolean {
    if (val === undefined || val === null || val === this.deps.appString.UNDEFINED || val.toString() === 'NaN') return true;

    const strVal = val.toString().trim();

    return strVal === '' || parseFloat(strVal) === 0;
  }

  // 🔹 Common Helper: Upload any file to Google Drive
  async uploadToDrive(fileBlob: Blob, fileName: string, mimeType: string): Promise<void> {
    const accessToken = localStorage.getItem('googleAccessToken');
    const googlEmail = localStorage.getItem('google_email');

    if (!accessToken) {
      console.error("No Google access token found");
      return;
    }

    try {
      // --- Step 1: Check if "edocsmart" folder exists ---
      let folderId: string | null = null;
      const searchRes = await fetch(
        "https://www.googleapis.com/drive/v3/files?q=name='edocsmart' and mimeType='application/vnd.google-apps.folder' and trashed=false",
        {
          method: "GET",
          headers: new Headers({ Authorization: `Bearer ${accessToken}` }),
        }
      );

      const searchData = await searchRes.json();
      if (searchData.files && searchData.files.length > 0) {
        folderId = searchData.files[0].id;
      } else {
        // --- Step 2: Create folder if not exists ---
        const folderMetadata = {
          name: "edocsmart",
          mimeType: "application/vnd.google-apps.folder",
        };
        const createRes = await fetch("https://www.googleapis.com/drive/v3/files", {
          method: "POST",
          headers: new Headers({
            Authorization: `Bearer ${accessToken}`,
            "Content-Type": "application/json",
          }),
          body: JSON.stringify(folderMetadata),
        });
        const createData = await createRes.json();
        folderId = createData.id;
      }

      // --- Step 3: Upload file ---
      const metadata = {
        name: fileName,
        mimeType,
        parents: [folderId],
      };

      const form = new FormData();
      form.append("metadata", new Blob([JSON.stringify(metadata)], { type: "application/json" }));
      form.append("file", fileBlob);

      const response = await fetch(
        "https://www.googleapis.com/upload/drive/v3/files?uploadType=multipart",
        {
          method: "POST",
          headers: new Headers({ Authorization: `Bearer ${accessToken}` }),
          body: form,
        }
      );

      const data = await response.json();
      this.deps.messageService.showSuccess(`Email - ${googlEmail} edocsmart/${fileName}`);
    } catch (err) {
    }
  }

  executeAction(
    action: string,
    docId: string,
    tableData: any[],
    partyName: string,
    invoiceNumber: string
  ) {
    switch (action) {
      case 'csv':
        this.downloadCSV(
          docId,
          tableData,
          partyName,
          invoiceNumber
        );
        break;
      case 'excel':
        this.downloadExcel(
          docId,
          tableData,
          partyName,
          invoiceNumber
        );
        break;
      case 'pdf':
        this.downloadPDF(
          docId,
          tableData,
          partyName,
          invoiceNumber
        );
        break;
    }
  }

}


