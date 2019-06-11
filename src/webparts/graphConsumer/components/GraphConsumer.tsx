import * as React from 'react';
import styles from './GraphConsumer.module.scss';
import * as strings from 'GraphConsumerWebPartStrings';
import { IGraphConsumerProps } from './IGraphConsumerProps';
import { IGraphConsumerState } from './IGraphConsumerState';
import { ClientMode } from './ClientMode';
import { IUserItem } from './IUserItem';
import { ILibrary } from './ILibrary';
import { IDocument } from './IDocument';
import { escape } from '@microsoft/sp-lodash-subset';
import { DefaultButton } from 'office-ui-fabric-react/lib/Button';
import { ContextualMenuItemType } from 'office-ui-fabric-react/lib/ContextualMenu';
import { ProgressIndicator } from 'office-ui-fabric-react/lib/ProgressIndicator';
import { Dialog, DialogType, DialogFooter } from 'office-ui-fabric-react/lib/Dialog';


import {
  autobind,
  PrimaryButton,
  TextField,
  Label,
  DetailsList,
  DetailsListLayoutMode,
  CheckboxVisibility,
  SelectionMode,
  Selection,
  MarqueeSelection,
  IColumn
} from 'office-ui-fabric-react';

import { AadHttpClient, MSGraphClient, GraphHttpClient, HttpClientResponse } from "@microsoft/sp-http";
import { mergeStyleSets } from 'office-ui-fabric-react/lib/Styling';

const classNames = mergeStyleSets({
  fileIconHeaderIcon: {
    padding: 0,
    fontSize: '16px'
  },
  fileIconCell: {
    textAlign: 'center',
    selectors: {
      '&:before': {
        content: '.',
        display: 'inline-block',
        verticalAlign: 'middle',
        height: '100%',
        width: '0px',
        visibility: 'hidden'
      }
    }
  },
  fileIconImg: {
    verticalAlign: 'middle',
    maxHeight: '16px',
    maxWidth: '16px'
  },
  controlWrapper: {
    display: 'flex',
    flexWrap: 'wrap'
  },
  exampleToggle: {
    display: 'inline-block',
    marginBottom: '10px',
    marginRight: '30px'
  },
  selectionDetails: {
    marginBottom: '20px'
  }
});
const controlStyles = {
  root: {
    margin: '0 30px 20px 0',
    maxWidth: '300px'
  }
};
const columns: IColumn[] = [
  {
    key: 'column1',
    name: 'File Type',
    //className: classNames.fileIconCell,
    //iconClassName: classNames.fileIconHeaderIcon,
    ariaLabel: 'Column operations for File type, Press to sort on File type',
    iconName: 'Page',
    isIconOnly: true,
    fieldName: 'name',
    minWidth: 16,
    maxWidth: 16,
    //onColumnClick: this._onColumnClick,
    onRender: (item: IDocument) => {
      return <img src={item.iconName} className={classNames.fileIconImg} alt={item.fileType + ' file icon'} />;
    }
  },
  {
    key: 'menu1',
    name: 'Nombre',
    minWidth: 210,
    maxWidth: 350,
    isResizable: true,
    //className: classNames.fileIconCell,
    //iconClassName: classNames.fileIconHeaderIcon,
    ariaLabel: 'Column operations for File type, Press to sort on File type',
    fieldName: 'name'
  },
  {
    key: 'column2',
    name: '',
    fieldName: 'menu',
    minWidth: 16,
    maxWidth: 16,
    isRowHeader: true,
    //onColumnClick: this._onColumnClick,
    data: 'string',
    isPadded: true,
    onRender: (item: IDocument) => {
      return (
        <div>
          <DefaultButton
            text="..."
            menuProps={{
              shouldFocusOnMount: true,
              items: [
                {
                  key: 'newItem',
                  name: 'Difusión de procedimiento',
                  onClick: () => console.log('New clicked'),
                  iconProps: {
                    iconName: 'Share'
                  }
                },
                {
                  key: 'rename',
                  name: 'Vincular documentos',
                  onClick: () => console.log('Rename clicked')
                }
              ]
            }}
          />
        </div>
      );
    }
  }, {
    key: 'column3',
    name: 'Fecha modificación',
    fieldName: 'createdBy',
    minWidth: 70,
    maxWidth: 90,
    isResizable: true,
    //onColumnClick: this._onColumnClick,
    onRender: (item: IDocument) => {
      return <span>{item.createdBy}</span>;
    },
    isPadded: true
  }, {
    key: 'column4',
    name: 'Descargable',
    fieldName: 'Descargable',
    minWidth: 70,
    maxWidth: 90,
    isResizable: true,
    //onColumnClick: this._onColumnClick,
    data: 'string',
    onRender: (item: IDocument) => {
      return item.descargable ? <a href={item.webUrl} title={item.name}>Descargar</a> : <span>No</span>;
    },
    isPadded: true
  }
  // {
  //   key: 'column4',
  //   name: 'Modified By',
  //   fieldName: 'modifiedBy',
  //   minWidth: 70,
  //   maxWidth: 90,
  //   isResizable: true,
  //   isCollapsible: true,
  //   data: 'string',
  //   onColumnClick: this._onColumnClick,
  //   onRender: (item: IDocument) => {
  //     return <span>{item.modifiedBy}</span>;
  //   },
  //   isPadded: true
  // },
  // {
  //   key: 'column5',
  //   name: 'File Size',
  //   fieldName: 'fileSizeRaw',
  //   minWidth: 70,
  //   maxWidth: 90,
  //   isResizable: true,
  //   isCollapsible: true,
  //   data: 'number',
  //   onColumnClick: this._onColumnClick,
  //   onRender: (item: IDocument) => {
  //     return <span>{item.fileSize}</span>;
  //   }
  // }
];

export default class GraphConsumer extends React.Component<IGraphConsumerProps, IGraphConsumerState> {
  private _selection: Selection;
  private _allItems: IDocument[];
  constructor(props: IGraphConsumerProps, state: IGraphConsumerState) {
    super(props);


    this._selection = new Selection({
      onSelectionChanged: () => {
        this.setState({
          selectionDetails: this._getSelectionDetails()
        });
      }
    });
    // Initialize the state of the component
    this.state = {
      users: [],
      searchFor: "",
      libraries: [],
      documents: [],
      selectionDetails: this._getSelectionDetails(),
      loadComplete: false,
      errorDialog: true,
      errorMessage: ""
    };


  }

  @autobind
  private _onSearchForChanged(newValue: string): void {

    // Update the component state accordingly to the current user's input
    this.setState({
      searchFor: newValue,
    });
  }

  private _getSearchForErrorMessage(value: string): string {
    // The search for text cannot contain spaces
    return (value == null || value.length == 0 || value.indexOf(" ") < 0)
      ? ''
      : `${strings.SearchForValidationErrorMessage}`;
  }

  private _getSelectionDetails(): string {
    // const selectionCount = this._selection.getSelectedCount();

    // switch (selectionCount) {
    //   case 0:
    //     return 'No items selected';
    //   case 1:
    //     return '1 item selected: ' + (this._selection.getSelection()[0] as IDocument).name;
    //   default:
    //     return `${selectionCount} items selected`;
    // }

    return "";
  }

  private _getLists(): void {

    

    this.props.context.graphHttpClient
      .get("v1.0/sites/imagen.sharepoint.com,5467ed13-ab7f-40fb-b5e4-e570adac0785,f7852e05-c564-4801-856a-5c88760f5ade/lists/cb6ef9de-cf38-4507-80e6-92e9409b7858/items?expand=Descargable",
        GraphHttpClient.configurations.v1)
      .then((resp: HttpClientResponse) => {
        var response = resp.json();
        console.log(response);
        //alert(res);
        response.then((res: any) => {

          var documents: Array<IDocument> = new Array<IDocument>();

          // Map the JSON response to the output array
          res.value.map((item: any) => {
            let icon = item.webUrl.split('.');
            let largo = icon.length;
            let docType = icon[largo - 1];

            let splitUrl = item.webUrl.split('/')
            let largoUrl = splitUrl.length;
            let nombre_documento = decodeURI(splitUrl[largoUrl - 1]);
            documents.push({
              webUrl: item.webUrl,
              createdBy: item.createdDateTime,
              id: item.id,
              fileType: docType,
              name: nombre_documento,
              descargable: item.fields.Descargable,
              menu: "",
              iconName: `https://static2.sharepointonline.com/files/fabric/assets/brand-icons/document/svg/${docType}_16x1.svg`
            });
          });

          // Update the component state accordingly to the result
          this.setState(
            {
              documents: documents,
              loadComplete: true,
            }
          );
        }).catch((e) => {
          console.log(e);
          //alert(e);
        });
      }).catch((err) => {
        console.log(err);
        //alert(err);
      });
  }

  private _getListsAAD(): void {
    this.props.context.aadHttpClientFactory
      .getClient('https://graph.microsoft.com')
      .then((client: AadHttpClient) => {
        // Search for the users with givenName, surname, or displayName equal to the searchFor value
        return client
          .get(
            `https://graph.microsoft.com/v1.0/sites/root`,
            AadHttpClient.configurations.v1
          );
      })
      .then(response => {
        return response.json();
      })
      .then(json => {
        alert("WebPart AAD " + json);
      })
      .catch(error => {
        alert("WebPart AAD " + error);
      });
  }
  @autobind
  private _search(): void {

    console.log(this.props.clientMode);

    // Based on the clientMode value search users
    switch (this.props.clientMode) {
      case ClientMode.aad:
        this._searchWithAad();
        break;
      case ClientMode.graph:
        this._searchWithGraph();
        break;
    }
  }

  private _searchWithAad(): void {

    // Log the current operation
    console.log("Using _searchWithAad() method");

    // Using Graph here, but any 1st or 3rd party REST API that requires Azure AD auth can be used here.
    this.props.context.aadHttpClientFactory
      .getClient('https://graph.microsoft.com')
      .then((client: AadHttpClient) => {
        // Search for the users with givenName, surname, or displayName equal to the searchFor value
        return client
          .get(
            `https://graph.microsoft.com/v1.0/users?$select=displayName,mail,userPrincipalName&$filter=(givenName%20eq%20'${escape(this.state.searchFor)}')%20or%20(surname%20eq%20'${escape(this.state.searchFor)}')%20or%20(displayName%20eq%20'${escape(this.state.searchFor)}')`,
            AadHttpClient.configurations.v1
          );
      })
      .then(response => {
        return response.json();
      })
      .then(json => {

        // Prepare the output array
        var users: Array<IUserItem> = new Array<IUserItem>();

        // Log the result in the console for testing purposes
        console.log(json);

        // Map the JSON response to the output array
        json.value.map((item: any) => {
          users.push({
            displayName: item.displayName,
            mail: item.mail,
            userPrincipalName: item.userPrincipalName,
          });
        });

        // Update the component state accordingly to the result
        this.setState(
          {
            users: users,
          }
        );
      })
      .catch(error => {
        console.error(error);
      });
  }

  private _searchWithGraph(): void {

    // Log the current operation
    console.log("Using _searchWithGraph() method");

    this.props.context.msGraphClientFactory
      .getClient()
      .then((client: MSGraphClient): void => {
        // From https://github.com/microsoftgraph/msgraph-sdk-javascript sample
        client
          .api("users")
          .version("v1.0")
          .select("displayName,mail,userPrincipalName")
          .filter(`(givenName eq '${escape(this.state.searchFor)}') or (surname eq '${escape(this.state.searchFor)}') or (displayName eq '${escape(this.state.searchFor)}')`)
          .get((err, res) => {

            if (err) {
              console.error(err);
              return;
            }

            // Prepare the output array
            var users: Array<IUserItem> = new Array<IUserItem>();

            // Map the JSON response to the output array
            res.value.map((item: any) => {
              users.push({
                displayName: item.displayName,
                mail: item.mail,
                userPrincipalName: item.userPrincipalName,
              });
            });

            // Update the component state accordingly to the result
            this.setState(
              {
                users: users,
              }
            );
          });
      });
  }

  componentDidMount() {
    this._getLists();
    //this._getListsAAD();
  }

  private _closeDialog = (): void => {
    this.setState({ loadComplete: true });
  };

  public render(): React.ReactElement<IGraphConsumerProps> {
    // Configure the columns for the DetailsList component
    let _usersListColumns = [
      {
        key: 'displayName',
        name: 'Display name',
        fieldName: 'displayName',
        minWidth: 50,
        maxWidth: 100,
        isResizable: true
      },
      {
        key: 'mail',
        name: 'Mail',
        fieldName: 'mail',
        minWidth: 50,
        maxWidth: 100,
        isResizable: true
      },
      {
        key: 'userPrincipalName',
        name: 'User Principal Name',
        fieldName: 'userPrincipalName',
        minWidth: 100,
        maxWidth: 200,
        isResizable: true
      },
    ];
    const { clientMode } = this.props;
    return (
      <div>
        <MarqueeSelection selection={this._selection}>
          <DetailsList
            items={this.state.documents}
            compact={true}
            columns={columns}
            selectionMode={SelectionMode.none}
            setKey="set"
            layoutMode={DetailsListLayoutMode.justified}
            isHeaderVisible={true}
            selection={this._selection}
            selectionPreservedOnEmptyClick={true}
            //onItemInvoked={this._onItemInvoked}
            enterModalSelectionOnTouch={true}
            ariaLabelForSelectionColumn="Toggle selection"
            ariaLabelForSelectAllCheckbox="Toggle selection for all items"
          />
        </MarqueeSelection>
        <Dialog
          hidden={this.state.loadComplete}
          onDismiss={this._closeDialog}
          dialogContentProps={{
            type: DialogType.normal,
            title: 'Cargando documentos, por favor, espere.',
          }}
          isBlocking={true}>
          <ProgressIndicator />
        </Dialog>

        <Dialog
          hidden={this.state.errorDialog}
          onDismiss={this._closeDialog}
          dialogContentProps={{
            type: DialogType.normal,
            title: 'Error al cargar la información',
          }}
          isBlocking={true}>
          <div>
            {this.state.errorMessage}
          </div>
          <DialogFooter>
            <PrimaryButton onClick={() => this.setState({ errorDialog: true })} text="Entendido" />
          </DialogFooter>
        </Dialog>

      </div>
    );
  }
}


/* Obtener usuarios
public render(): React.ReactElement<IGraphConsumerProps> {
    // Configure the columns for the DetailsList component
    let _usersListColumns = [
      {
        key: 'displayName',
        name: 'Display name',
        fieldName: 'displayName',
        minWidth: 50,
        maxWidth: 100,
        isResizable: true
      },
      {
        key: 'mail',
        name: 'Mail',
        fieldName: 'mail',
        minWidth: 50,
        maxWidth: 100,
        isResizable: true
      },
      {
        key: 'userPrincipalName',
        name: 'User Principal Name',
        fieldName: 'userPrincipalName',
        minWidth: 100,
        maxWidth: 200,
        isResizable: true
      },
    ];
    const { clientMode } = this.props;
    return (
      <div className={styles.graphConsumer}>
        <div className={styles.container}>
          <div className={styles.row}>
            <div className={styles.column}>
              <span className={styles.title}>Search for a user!</span>
              <p className={styles.form}>
                <TextField
                  label={strings.SearchFor}
                  required={true}
                  value={this.state.searchFor}
                  onChanged={this._onSearchForChanged}
                  onGetErrorMessage={this._getSearchForErrorMessage}
                />
              </p>
              {
                (clientMode === ClientMode.aad || clientMode === ClientMode.graph) ?
                  <p className={styles.form}>
                    <PrimaryButton
                      text='Search'
                      title='Search'
                      onClick={this._search}
                    />
                  </p>
                  : <p>Configure client mode by editing web part properties.</p>
              }
              {
                (this.state.users != null && this.state.users.length > 0) ?
                  <p className={styles.form}>
                    <DetailsList
                      items={this.state.users}
                      columns={_usersListColumns}
                      setKey='set'
                      checkboxVisibility={CheckboxVisibility.hidden}
                      selectionMode={SelectionMode.none}
                      layoutMode={DetailsListLayoutMode.fixedColumns}
                      compact={true}
                    />
                  </p>
                  : null
              }
            </div>
          </div>
        </div>
      </div>
    );
  }




  this.props.context.msGraphClientFactory
      .getClient()
      .then((client: MSGraphClient): void => {
        client
          //.api("/sites/imagen.sharepoint.com,5467ed13-ab7f-40fb-b5e4-e570adac0785,f7852e05-c564-4801-856a-5c88760f5ade/lists/cb6ef9de-cf38-4507-80e6-92e9409b7858/items?expand=Descargable")
          .api("sites/root")
          .version("v1.0")
          //.select("id,webUrl,createdBy,lastModifiedBy,createdDateTime,Descargable")
          //.expand("Descargable")
          //.filter(`(givenName eq '${escape(this.state.searchFor)}') or (surname eq '${escape(this.state.searchFor)}') or (displayName eq '${escape(this.state.searchFor)}')`)
          .get((err, res) => {
            if (err) {
              //console.error(err);
              alert("WebPart " + err.message)
              this.setState({ errorDialog: false, loadComplete: true, errorMessage: err.message + " " + err.statusCode })
              return;
            }

            alert("WebPart " + res);

            // Prepare the output array


            //console.log(documents);
          });
      });
*/