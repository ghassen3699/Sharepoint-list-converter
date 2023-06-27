import * as React from 'react';
// import styles from './SharePointListToExcelFile.module.scss';
import { ISharePointListToExcelFileProps } from './ISharePointListToExcelFileProps';
import { PrimaryButton } from 'office-ui-fabric-react';
import * as ExcelJS from 'exceljs';
import "@pnp/sp/items";
import "@pnp/sp/webs";
import "@pnp/sp/lists";
import { Web } from '@pnp/sp/webs';
// import { Web } from '@pnp/sp/webs';


export default class SharePointListToExcelFile extends React.Component<ISharePointListToExcelFileProps, {}> {
  // State of webpart 
  public state = {
    listData : [] as any 
  };   

  // Get list data from SharePoint
  getListData = async() => {
    const data = await Web(this.props.url).lists.getByTitle("les demandes").items();
    this.setState({listData:data})
  }

  // add sharepoint items to Excel file
  convertListDataToExcelFile = () => {
    const workbook = new ExcelJS.Workbook();
    const worksheet = workbook.addWorksheet('List Data');
    const fileName = 'SharePointList.xlsx';
    const columns = ['Raison', 'Nombredejour', 'Status'];
  
    worksheet.addRow(columns); // Add header row
  
    this.state.listData.forEach((item: any) => {
      const rowData = columns.map(column => item[column]);
      worksheet.addRow(rowData); // Add data rows
    });
  
    worksheet.columns.forEach(column => {
      let maxColumnWidth = 0;
      column.eachCell({ includeEmpty: true }, cell => {
        const columnWidth = cell.value ? cell.value.toString().length + 2 : 10;
        if (columnWidth > maxColumnWidth) {
          maxColumnWidth = columnWidth;
        }
      });
      column.width = maxColumnWidth;
    });

    worksheet.getRow(1).eachCell(cell => {
      cell.fill = {
        type: 'pattern',
        pattern: 'solid',
        fgColor: { argb: '000000' }, // Black color code
      };
      cell.font = {
        color: { argb: 'FFFFFF' }, // White color code
        bold: true,
      };
    });
  
    workbook.xlsx.writeBuffer().then((buffer: ArrayBuffer) => {
      const blob = new Blob([buffer], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
  
      const url = URL.createObjectURL(blob);
      const link = document.createElement('a');
      link.href = url;
      link.download = fileName;
      link.click();
    });
  }

  componentDidMount(): void {
    this.getListData()
  }

  public render(): React.ReactElement<ISharePointListToExcelFileProps> {
    
    return (
      <div>
          Download Excel File
          <br></br>
          <PrimaryButton onClick={() => this.convertListDataToExcelFile()}>Click</PrimaryButton>
      </div>
    );
  }
}
