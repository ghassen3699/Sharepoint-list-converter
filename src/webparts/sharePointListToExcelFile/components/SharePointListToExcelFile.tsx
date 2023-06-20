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

  convertListDataToExcelFile = () => {
    const workbook = new ExcelJS.Workbook();
    const worksheet = workbook.addWorksheet('List Data');
    const fileName = 'SharePointList.xlsx';
    worksheet.addRow(['Raison', 'Nombredejour', 'Status']); // Add header row

    this.state.listData.forEach((item : any) => {
      worksheet.addRow([item.Raison, item.Nombredejour, item.Status]); // Add data rows
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
          webpart
          <PrimaryButton onClick={() => this.convertListDataToExcelFile()}>Click</PrimaryButton>
      </div>
    );
  }
}
