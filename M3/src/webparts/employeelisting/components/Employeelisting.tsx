import * as React from 'react';
// import styles from './Employeelisting.module.scss';
import { IEmployeelistingProps } from './IEmployeelistingProps';
// import { escape } from '@microsoft/sp-lodash-subset';
// import { sp } from "@pnp/sp/presets/all"; 
import { Dropdown, IDropdownOption } from '@fluentui/react/lib/Dropdown';
import { Announced } from '@fluentui/react/lib/Announced';
import { TextField, ITextFieldStyles } from '@fluentui/react/lib/TextField';
import { DetailsList, DetailsListLayoutMode, Selection, IColumn } from '@fluentui/react/lib/DetailsList';
import { MarqueeSelection } from '@fluentui/react/lib/MarqueeSelection';
import { mergeStyles } from '@fluentui/react/lib/Styling';
// import { Text } from '@fluentui/react/lib/Text';
import { sp } from "@pnp/sp/presets/all";
import "@pnp/sp/webs"
import "@pnp/sp/lists";
import "@pnp/sp/items";
import { DatePicker, Fabric, Icon } from 'office-ui-fabric-react';
import { Dialog, DialogFooter } from '@fluentui/react/lib/Dialog'; // DialogFooter
import { PrimaryButton, DefaultButton } from '@fluentui/react/lib/Button'; // IButtonProps
import { PeoplePicker, PrincipalType } from "@pnp/spfx-controls-react/lib/PeoplePicker";
// import { Fabric } from 'office-ui-fabric-react';
// import { Fabric } from 'office-ui-fabric-react/lib/Fabric';
// import { format } from "date-fns";


// export default class Employeelisting extends React.Component<IEmployeelistingProps, {}> {

// public render(): React.ReactElement<IEmployeelistingProps> {
//   const {
//     description,
//     isDarkTheme,
//     environmentMessage,
//     hasTeamsContext,
//     userDisplayName
//   } = this.props;

//     return (
//       <div>Hello</div>
//     );
//   }
// }


const exampleChildClass = mergeStyles({
  display: 'block',
  marginBottom: '10px',
});



const textFieldStyles: Partial<ITextFieldStyles> = { root: { maxWidth: '300px' } };

export interface IDetailsListBasicExampleItem {
  ID: number;
  Name: string;
  DOB: string;
  Experience: number;
  DeptName: string;
  DeptNameId: any;
  Manager: any[];
  ManagerId: any;

}

export interface IDetailsListBasicExampleState {
  items: IDetailsListBasicExampleItem[];
  selectionDetails: string;
  announcedMessage: any;
  FilterData: IDetailsListBasicExampleItem[];
  ItemId: any;
  DepartmentId: any;
  id: any;
  Name: any;
  DOB: any;
  Experience: any;
  selectedusers: string[];
  Manager: [];
  ManagerId: any;
  SelectedItem: any;
  SelectedItemup: any;
  HideDialogup: boolean;
  HideDialogconfirmation: boolean;
  HideDialog: boolean;
  projectlookupvalues: IDropdownOption[];
  DeptNameId: any;
  DeptName: any;
  pluser: any,
  EditMode: boolean,
  isUserInGroup: boolean,
  hidebutton: boolean,
  UserEmail: any,
  gusers: any,
  checkFields: boolean,
  
}


export default class Employeelisting extends React.Component<IEmployeelistingProps, IDetailsListBasicExampleState> {
  private _selection: Selection;
  private _columns: IColumn[];


  constructor(props: IEmployeelistingProps) {
    super(props);
    sp.setup({
      spfxContext: this.props.spfxcontext
    });

    this.state = {
      items: [],
      selectionDetails: '',
      announcedMessage: undefined,
      DepartmentId: '',
      FilterData: [],
      ItemId: [],
      projectlookupvalues: [],
      SelectedItemup: '',
      selectedusers: [],
      DeptNameId: '',
      HideDialogup: true,
      HideDialogconfirmation: true,
      SelectedItem: undefined,
      pluser: [],
      id: '',
      Name: '',
      DOB: '',
      Experience: '',
      DeptName: '',
      Manager: [],
      ManagerId: '',
      HideDialog: true,
      EditMode: false,
      isUserInGroup: false,
      hidebutton: true,
      UserEmail: [],
      gusers: [],
      checkFields: true,
      



    };

    this._selection = new Selection({
      onSelectionChanged: () => this.setState({ selectionDetails: this._getSelectionDetails() }),
    });


    this._columns = [
      { key: 'column1', name: 'Action', fieldName: 'Action', minWidth: 100, maxWidth: 100, isResizable: true },
      { key: 'column2', name: 'ID', fieldName: 'ID', minWidth: 100, maxWidth: 100, isResizable: true },
      { key: 'column3', name: 'Name', fieldName: 'Name', minWidth: 100, maxWidth: 200, isResizable: true, onColumnClick: this._onColumnClick }, // For sorting
      { key: 'column4', name: 'DOB', fieldName: 'DOB', minWidth: 100, maxWidth: 200, isResizable: true },
      { key: 'column5', name: 'Experience', fieldName: 'Experience', minWidth: 100, maxWidth: 200, isResizable: true },
      { key: 'column6', name: 'DeptName', fieldName: 'DeptName', minWidth: 100, maxWidth: 200, isResizable: true },
      { key: 'column7', name: 'Manager', fieldName: 'Manager', minWidth: 100, maxWidth: 200, isResizable: true },
    ];
  }

  render(): React.ReactElement<IEmployeelistingProps> {
    const { items, selectionDetails } = this.state;

    return (
      <div>
        <Fabric>
          <div className={exampleChildClass}>{selectionDetails}</div>
          <Announced message={selectionDetails} />
          <TextField
            className={exampleChildClass}
            label="Filter by Name:"
            onChange={this._onFilter}
            styles={textFieldStyles}
          />

          <Announced message={`Number of items after filter applied: ${items.length}.`} />{
            this.state.hidebutton == false &&
            <PrimaryButton text="Add Employee" onClick={() => { this.setState({ HideDialog: false }), this.reset() }} />}
          <MarqueeSelection selection={this._selection}>
            <DetailsList
              items={this.state.items}
              columns={this._columns}
              setKey="set"
              layoutMode={DetailsListLayoutMode.justified}
              selection={this._selection}
              selectionPreservedOnEmptyClick={true}
              ariaLabelForSelectionColumn="Toggle selection"
              ariaLabelForSelectAllCheckbox="Toggle selection for all items"
              checkButtonAriaLabel="select row"
              onItemInvoked={this._onItemInvoked}
            />
          </MarqueeSelection>
        </Fabric>
        <div className='adddialog'>
          <Dialog
            hidden={this.state.HideDialog}>
            <h1>Add Employee</h1>
            <div className='table'>
              <table>
                <tr className='Name'>
                  <td>
                    Name :
                  </td>
                  <td>
                    <input type="text" id="Name" value={this.state.Name ? this.state.Name : ''} onChange={(e) => { this.setState({ Name: e.target.value }); }} />
                    {this.state.checkFields == false && (<p style={{ color: 'red' }}>Invalid Name</p>)}
                  </td>
                </tr>
                <tr className='DOB'>
                  <td>
                    DOB :
                  </td>
                  <td>
                    <input type="date" id="DOB" value={this.state.DOB ? this.state.DOB : ''} onChange={(e) => { this.setState({ DOB: e.target.value }); }} />
                  </td>
                </tr>
                <tr className='Experience'>
                  <td>
                    Experience :
                  </td>
                  <td>
                    <input type="number" min='0' id="Experience" value={this.state.Experience ? this.state.Experience : ''} onChange={(e) => { this.setState({ Experience: e.target.value }); }} />
                  </td>
                </tr>
                <tr className='Department'>
                  <td>
                    Department :
                  </td>
                  <td>
                    <Dropdown placeholder="Select a Department" options={this.state.projectlookupvalues} onChange={(e, val) => { this.onDropdownchange(e, val) }} ></Dropdown>
                  </td>
                </tr>
                <tr className='Manager'>
                  <td>
                    Manager :
                  </td>
                  <td>
                    <PeoplePicker
                      context={this.props.spfxcontext}
                      personSelectionLimit={5}
                      showtooltip={true}
                      required={true}
                      disabled={false}
                      onChange={this._getPeoplePicker}
                      showHiddenInUI={false}
                      ensureUser={true}
                      principalTypes={[PrincipalType.User]}
                      resolveDelay={1000}/>
                  </td>
                </tr>
              </table>
            </div>
            <DialogFooter>
            {
             (this.state.Name != '' && this.state.DOB !=null && this.state.Experience !='' && this.state.SelectedItemup !=0 && this.state.pluser.length>0 ) &&
             ( <PrimaryButton text="Save" onClick={() => { this.onChangeSetName()}} /> )}
              <DefaultButton text="Cancel" onClick={() => { this.setState({ HideDialog: true }) }} />
            </DialogFooter>
          </Dialog>
          {/* ---------------------------Update info Dialog --------------------------------- */}
          <Dialog
            hidden={this.state.HideDialogup}>
            <h1>Update Info</h1>
            <div className='table'>
              <table>
                <tr className='Name'>
                  <td>
                    Name :
                  </td>
                  <td>
                    <input type="text" id="Name" value={this.state.Name ? this.state.Name : ''} onChange={(e) => { this.setState({ Name: e.target.value }); }} />
                    {this.state.checkFields == false && (<p style={{ color: 'red' }}>Invalid Name</p>)}
                  </td>
                </tr>
                <tr className='DOB'>
                  {/* <td>
                    DOB :
                  </td>
                  <td>
                    <input type="date" id="DOB" value={this.state.DOB ? this.state.DOB : ''} onChange={(e) => { this.setState({ DOB: e.target.value }); }} />
                  </td> */}
                  <td>
                    DOB:
                  </td>
                  <td><DatePicker id="DOB"
                    value={new Date(this.state.DOB)}
                    onSelectDate={(selectedDate) => {
                      this.setState({ DOB: selectedDate });
                    }} isRequired/>
                  </td>
                </tr>
                <tr className='Experience'>
                  <td>
                    Experience :
                  </td>
                  <td>
                    <input type="number" min='0' id="Experience" value={this.state.Experience ? this.state.Experience : ''} onChange={(e) => { this.setState({ Experience: e.target.value }); }} />
                  </td>
                </tr>
                <tr className='Departmentup'>
                  <td>
                    Department :
                  </td>
                  <td>
                    <Dropdown placeholder="Select a Department" options={this.state.projectlookupvalues} defaultSelectedKey={this.state.SelectedItemup} onChange={(e, val) => { this.onDropdownchange(e, val) }} ></Dropdown>
                  </td>
                </tr>
                <tr className='Manager'>
                  <td>
                    Manager :
                  </td>
                  <td>
                    <PeoplePicker
                      context={this.props.spfxcontext}
                      personSelectionLimit={5}
                      showtooltip={true}
                      required={true}
                      disabled={false}
                      onChange={this._getPeoplePicker}
                      defaultSelectedUsers={this.state.selectedusers}
                      showHiddenInUI={false}
                      ensureUser={true}
                      principalTypes={[PrincipalType.User]}
                      resolveDelay={1000} />
                  </td>
                </tr>
              </table>
            </div>
            <DialogFooter> {
             (this.state.Name != '' && this.state.DOB !=null && this.state.Experience !='' && this.state.SelectedItemup !=0 && this.state.pluser.length>0 ) &&
             (<PrimaryButton text="Update" onClick={() => { this.onChangeSetNameUp() }} />)}
              <DefaultButton text="Cancel" onClick={() => { this.setState({ HideDialogup: true }) }} />
            </DialogFooter>
          </Dialog>
          {/* ---------------------------Yes or No Dialog--------------------------------- */}
          <Dialog hidden={this.state.HideDialogconfirmation} >
            <text>
              “Are you sure, you want to update the details?”
            </text>
            <DialogFooter>
              <PrimaryButton text="Yes" onClick={() => { this.UpdateItem(this.state.ItemId) }} />
              <DefaultButton text="No" onClick={() => { this.setState({ HideDialogconfirmation: true }) }} />
            </DialogFooter>
          </Dialog>
        </div>
      </div>
    );
  }

  // function for hide save button
public hideBtn = () =>{
  if(this.state.Name != '' && this.state.DOB !=null && this.state.Experience !='' && this.state.SelectedItemup !=0 && this.state.pluser.length >0){
    this.state.hidebutton == false;
  }
  else{
    this.state.hidebutton == true;
  }
}
//function for hide update button
public hideUpBtn = () =>{
  if(this.state.Name != '' && this.state.DOB !=null && this.state.Experience !='' && this.state.SelectedItemup !=0 && this.state.pluser.length >0){
    this.state.hidebutton == false;
  }
  else{
    this.state.hidebutton == true;
  }
}

  // validation---------------------------------------------------------------
  // val for createItem
  public onChangeSetName = () => {
    let pattern = new RegExp("^[A-Za-z0-9 ]+$"); // ^[a-zA-Z0-9.-]*$
    let isValid = pattern.test(this.state.Name);
    if (isValid) {
      this.setState({ checkFields: true });
      this.hideBtn();
      this.createItem();
    } else {
      this.setState({ checkFields: false }, () => { });
    }
  };

   // val for UpdateItem..................................................
  public onChangeSetNameUp = () => {
    let pattern = new RegExp("^[A-Za-z0-9 ]+$");
    let isValid = pattern.test(this.state.Name);
    if (isValid) {
      this.setState({ checkFields: true });
      this.hideUpBtn();
      this.updatedialog();
    } else {
      this.setState({ checkFields: false }, () => { });
    }
  };



  public reset = async () => {
    this.setState({ Name: '', DOB: null, Experience: '', SelectedItemup: 0, selectedusers: [], })
  }
  public componentDidMount = () => {
    this.getListItems();
    this.checkUserInGroup();
    this._getcurrentuser();
    this._getdeplookupfield();

    // this.getListItem();
  }

  public getListItems = async () => {

    await sp.web.lists.getByTitle("Employee").items.select("ID", "Name", "DOB", "FieldValuesAsText/DOB", "Experience", "DeptName/ID", "DeptName/DepartmentName", "Manager/ID", "Manager/EMail").expand("FieldValuesAsText", "DeptName", "Manager").get().then(items => {

      let AllData: { Action: any; ID: any; Name: any; DOB: any; Experience: number; DeptName: any; DeptNameId: any; Manager: any; ManagerId: any; }[] = [];
      items.map((data) => {

        let Allusers: any[] = [];
        data.Manager.map((val: any) => {
          Allusers.push(val.EMail)
        })

        AllData.push({
          ID: data.ID,
          Name: data.Name,
          DOB: data["FieldValuesAsText"].DOB,
          Experience: data.Experience,
          DeptName: data.DeptName.DepartmentName,
          DeptNameId: data.DeptName.ID,
          Manager: Allusers,
          ManagerId: data.Manager.ID,
          Action: (
            <div>
              <Icon
                iconName='delete' onClick={() => { this.deleteItem(data.ID) }} style={{ marginRight: 30 }}>
              </Icon>
              <Icon
                iconName='edit' onClick={() => { this.editModeItems(data.ID) }} >
              </Icon>
            </div>
          )
        })
        console.dir(items);

      })
      this.setState({
        items: AllData,
        selectionDetails: this._getSelectionDetails(),
        FilterData: AllData,
      });
      // console.log(data);
    }).catch((e) => {
      console.log(e);
    })
  }


  private _getSelectionDetails(): string {
    const selectionCount = this._selection.getSelectedCount();


    switch (selectionCount) {
      case 0:
        return 'No items selected';
      case 1:
        return '1 item selected: ' + (this._selection.getSelection()[0] as IDetailsListBasicExampleItem).Name;
      default:
        return `${selectionCount} items selected`;
    }
  }
  private _onItemInvoked = (item: IDetailsListBasicExampleItem): void => {
    alert(`Item invoked: ${item.Name}`);
  };


  private _onFilter = (ev: React.FormEvent<HTMLInputElement | HTMLTextAreaElement>, text: string): void => {
    this.setState({
      items: text ? this.state.FilterData.filter(i => i.Name.toLowerCase().indexOf(text) > -1) : this.state.FilterData,
    });
  };





  //-----------for email ---------------------------------
  public sendMail = async () => {
    let addressString: string = await sp.utility.getCurrentUserEmailAddresses();
    await sp.utility.sendEmail({
      To: [addressString],
      Subject: "This email is about...",
      Body: "<b>New Item is Added....!!</b>",
      AdditionalHeaders: {
        "content-type": "text/html"
      },
    });
  }

  // public getAllUsers  = async () => {
  //   const groupName = "AdminGroup";
  //   sp.web.siteGroups
  //     .getByName(groupName)
  //     .users()
  //     .then((users) => {
  //       console.log(users);
  //     })
  //     .catch((error) => {
  //       console.log(error);
  //     });
  // }
  //   const groupID = 16;
  //   const users = await sp.web.siteGroups.getById(groupID).users();
  // // const group = await sp.web.siteGroups.getByName("AdminGroup");
  // // const users = await group.users();
  // console.log(users);
  //  /^[A-Za-z]+$/; ---> for name

  //----------------hide and show add emlpoyee button------------------
  //----------------check the user in the group------------------------
  public checkUserInGroup = async () => {
    const groupID = 16;
    const users = await sp.web.siteGroups.getById(groupID).users();
    users.forEach((item) => {
      if (item.Email.toLowerCase() == this.state.UserEmail.toLowerCase()) {
        this.setState({ hidebutton: false })
      }
    })
    this.setState({
      gusers: users,
    })
    console.log(users);
  }
  //-----------------fetch current user with group users------------- 
  public async _getcurrentuser(): Promise<any> {
    const currentuser = await sp.profiles.userProfile;

    this.setState({
      UserEmail: currentuser.SipAddress,
    }, () => { this.checkUserInGroup() })
    console.log(currentuser);
  }


  //--------------for dropdown-----------------------
  public _getdeplookupfield = async () => {
    const allItems: any[] = await sp.web.lists.getByTitle("Department").items.getAll();
    let dropdowndep: IDropdownOption[] = [];
    allItems.forEach(Department => {
      dropdowndep.push({ key: Department.ID, text: Department.DepartmentName });
    })
    this.setState({
      projectlookupvalues: dropdowndep
    });
  }

  public onDropdownchange(event: React.FormEvent<HTMLDivElement>, item: IDropdownOption) {
    console.log();
    this.setState({ SelectedItem: item.key, SelectedItemup: item.key })

  }


  //Create Item...........................
  public createItem = async () => {

    sp.web.lists.getByTitle("Employee")
      .items.add({
        Name: this.state.Name,
        DOB: this.state.DOB,
        Experience: this.state.Experience,
        DeptNameId: this.state.SelectedItem,
        ManagerId: { results: this.state.pluser }
      }).then(() => {
        this.setState({ HideDialog: true })
        this.getListItems();
        this.sendMail();
      }).catch((err) => {
        console.log(err);
      });
  }

  //  edit items .........................................
  public editModeItems = async (Id: any) => {
    let editItem = this.state.items.filter((x: any) => { return x.ID == Id; })[0];
    this.setState({
      Name: editItem.Name,
      DOB: editItem.DOB,
      Experience: editItem.Experience,
      ItemId: editItem.ID,
      DeptName: editItem.DeptName,
      SelectedItemup: editItem.DeptNameId,
      selectedusers: editItem.Manager,
      ManagerId: { results: this.state.pluser },
      // ManagerId: editItem.ManagerId,
      HideDialogup: false,
      EditMode: true,
    });
  }

  // --------------Yes Or No Dialog --------------------------------
  public updatedialog = async () => {
    this.setState({
      HideDialogconfirmation: false,
    })
  }
  //update Item
  public UpdateItem = async (ItemId: any) => {
    await sp.web.lists.getByTitle("Employee").items.getById(ItemId)
      .update({
        Name: this.state.Name,
        DOB: this.state.DOB,
        Experience: this.state.Experience,
        DeptNameId: this.state.SelectedItem,
        ManagerId: { results: this.state.pluser }
        // ManagerId: this.state.ManagerId,
      }).then((data) => {
        this.setState({ HideDialogconfirmation: true, HideDialogup: true });
        this.reset(),
          this.getListItems();
        console.log(data);
      })
      .catch((err) => {
        console.log(err);
      });
  }

  // delete item................
  public deleteItem = async (ID: any) => {
    console.log(ID);
    await sp.web.lists.getByTitle("Employee").items.getById(ID).delete().then((data) => {
      console.log(data);
      this.getListItems();
    }).catch((err) => {
      console.log(err);
    });
  }


  // ---------peoplepicker-------------------
  private _getPeoplePicker = (pluser: any) => {
    let AllManager: any[] = [];
    pluser.map((val: any) => {
      AllManager.push(val.id)
    })
    this.setState({ pluser: AllManager });
  }


  // For ascending and descending order on Name Column 
  private _onColumnClick = (ev: React.MouseEvent<HTMLElement>, column: IColumn): void => {
    const { items } = this.state;
    const newColumns: IColumn[] = this._columns.slice();
    const currColumn: IColumn = newColumns.filter(currCol => column.key === currCol.key)[0];
    newColumns.forEach((newCol: IColumn) => {
      if (newCol === currColumn) {
        currColumn.isSortedDescending = !currColumn.isSortedDescending;
        currColumn.isSorted = true;
        this.setState({
          announcedMessage: `${currColumn.name} is sorted ${currColumn.isSortedDescending ? 'descending' : 'ascending'
            }`,
        });
      } else {
        newCol.isSorted = false;
        newCol.isSortedDescending = true;
      }
    });
    const newItems = _copyAndSort(items, currColumn.fieldName!, currColumn.isSortedDescending);
    this.setState({
      items: newItems,
    });
  };
}

function _copyAndSort<T>(items: T[], columnKey: string, isSortedDescending?: boolean): T[] {
  const key = columnKey as keyof T;
  return items.slice(0).sort((a: T, b: T) => ((isSortedDescending ? a[key] < b[key] : a[key] > b[key]) ? 1 : -1));
}






