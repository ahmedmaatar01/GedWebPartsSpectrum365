import * as React from 'react';
import { IDropdownOption, IChoiceGroupOption, Dropdown } from 'office-ui-fabric-react';
import styles from './FormulaireDeSoumessionDesDocuments.module.scss';
import type { IFormulaireDeSoumessionDesDocumentsProps } from './IFormulaireDeSoumessionDesDocumentsProps';
import type { IFormulaireDeSoumessionDesDocumentsState } from './IFormulaireDeSoumessionDesDocumentsState';
import { SPOperations, SPListColumn } from '../../Services/SpFormulaireSoumessionService';
import ButtonGrid from './ButtonGrid';
import CreateDocumentModal from './CreateDocumentModal';
import UploadDocumentModal from './UploadDocumentModal';
import AddMetadataModal from './AddMetadataModal';
import CreateFolderModal from './CreateFolderModal';

const metadataTypeOptions: IDropdownOption[] = [
  { key: 'Text', text: 'Ligne de texte' },
  { key: 'Choice', text: 'Choix' },
  { key: 'Number', text: 'Nombre' },
  { key: 'Boolean', text: 'Oui/Non' },
  { key: 'Image', text: 'Image' },
];

const choiceGroupOptions: IChoiceGroupOption[] = [
  { key: 'Yes', text: 'Oui' },
  { key: 'No', text: 'Non' },
];

// interface IFormulaireDeSoumessionDesDocumentsState {
//   listColumns: SPListColumn[];
//   listTiltes: string[];
//   listItems: any[];
//   status: string;
//   Titre_list_item: string;
//   showModal: boolean;
//   listItemId: string;
//   selectedDocumentType: string;
//   metadata: { [key: string]: any };
//   uploadFile: File | null;
//   isUploadMode: boolean;
//   showCreateModal: boolean;
//   showUploadModal: boolean;
//   showAddMetadataModal: boolean;
//   newMetadataField: string;
//   newMetadataDescription: string;
//   newMetadataType: string;
//   choices: string[];
//   showCreateFolderModal: boolean;
//   newFolderName: string;
//   users: { id: string; title: string }[];
//   selectedUser: string | null;
// }

export default class FormulaireDeSoumessionDesDocuments extends React.Component<IFormulaireDeSoumessionDesDocumentsProps, IFormulaireDeSoumessionDesDocumentsState> {
  public _spOperations: SPOperations;

  constructor(props: IFormulaireDeSoumessionDesDocumentsProps) {
    super(props);
    this._spOperations = new SPOperations();
    this.state = {
      listColumns: [],
      listTiltes: [],
      listItems: [],
      status: '',
      Titre_list_item: '',
      showModal: false,
      listItemId: '',
      selectedDocumentType: 'txt',
      metadata: {},
      uploadFile: null,
      isUploadMode: false,
      showCreateModal: false,
      showUploadModal: false,
      showAddMetadataModal: false,
      newMetadataField: '',
      newMetadataDescription: '',
      newMetadataType: 'Text',
      choices: [''],
      showCreateFolderModal: false,
      newFolderName: '',
      users: [],
      selectedUser: null,
    };
  }

  public openCreateModal = () => {
    this.setState({ showCreateModal: true }, () => {
      if (this.props.list_title) {
        this._spOperations.GetListColumns(this.props.context, this.props.list_title)
          .then((results: SPListColumn[]) => {
            const filteredColumns = results.filter(column =>
              column.internalName !== 'ContentType' &&
              column.internalName !== 'Title' &&
              column.internalName !== '_ExtendedDescription'
            );
            const metadata: { [key: string]: any } = {};
            filteredColumns.forEach(column => {
              metadata[column.internalName] = '';
            });
            this.setState({ listColumns: filteredColumns, metadata });
          })
          .catch(error => {
            console.error('Error getting list columns:', error);
          });

        this._spOperations.GetUsers(this.props.context)
          .then((users) => {
            this.setState({ users });
          })
          .catch(error => {
            console.error('Error getting users:', error);
          });
      }
    });
  };

  private handleUserChange = (event: React.FormEvent<HTMLDivElement>, option?: IDropdownOption, internalName?: string) => {
    if (option && internalName) {
      this.setState(prevState => ({
        metadata: {
          ...prevState.metadata,
          [`${internalName}Id`]: option.key,
        },
      }));
    }
  };

  public openUploadModal = () => {
    this.setState({ showUploadModal: true }, () => {
      if (this.props.list_title) {
        this._spOperations.GetListColumns(this.props.context, this.props.list_title)
          .then((results: SPListColumn[]) => {
            const filteredColumns = results.filter(column =>
              column.internalName !== 'ContentType' &&
              column.internalName !== 'Title' &&
              column.internalName !== '_ExtendedDescription' &&
              column.internalName !== 'FileLeafRef'
            );
            const metadata: { [key: string]: any } = {};
            filteredColumns.forEach(column => {
              metadata[column.internalName] = '';
            });
            this.setState({
              listColumns: filteredColumns,
              metadata,
            });
          })
          .catch(error => {
            console.error('Error getting list columns:', error);
          });
      }
    });
  };

  public openAddMetadataModal = () => {
    this.setState({ showAddMetadataModal: true });
  };

  public openCreateFolderModal = () => {
    this.setState({ showCreateFolderModal: true }, () => {
      if (this.props.list_title) {
        this._spOperations.GetListColumns(this.props.context, this.props.list_title)
          .then((results: SPListColumn[]) => {
            const filteredColumns = results.filter(column =>
              column.internalName !== 'ContentType' &&
              column.internalName !== 'Title' &&
              column.internalName !== '_ExtendedDescription' &&
              column.internalName !== 'Nom_du_dossier'
            );
            const metadata: { [key: string]: any } = {};
            filteredColumns.forEach(column => {
              metadata[column.internalName] = '';
            });
            this.setState({ listColumns: filteredColumns, metadata });
          })
          .catch(error => {
            console.error('Error getting list columns:', error);
          });
      }
    });
  };

  private handleNewMetadataFieldChange = (event: React.ChangeEvent<HTMLInputElement>) => {
    this.setState({ newMetadataField: event.target.value });
  };

  private handleNewMetadataDescriptionChange = (event: React.ChangeEvent<HTMLInputElement>) => {
    this.setState({ newMetadataDescription: event.target.value });
  };

  private handleNewMetadataTypeChange = (event: React.FormEvent<HTMLDivElement>, option?: IDropdownOption) => {
    if (option) {
      this.setState({ newMetadataType: option.key as string, choices: [''] });
    }
  };

  private handleInputChange = (internalName: string) => (event: React.ChangeEvent<HTMLInputElement>) => {
    const value = event.target.type === 'checkbox' ? event.target.checked : event.target.value;
    this.setState(prevState => ({
      metadata: {
        ...prevState.metadata,
        [internalName]: value,
      },
    }));
  };

  private handleSelectChange = (internalName: string) => (event: React.ChangeEvent<HTMLSelectElement>) => {
    const value = event.target.value;
    this.setState(prevState => ({
      metadata: {
        ...prevState.metadata,
        [internalName]: value,
      },
    }));
  };

  private handleTextareaChange = (internalName: string) => (event: React.ChangeEvent<HTMLTextAreaElement>) => {
    const value = event.target.value;
    this.setState(prevState => ({
      metadata: {
        ...prevState.metadata,
        [internalName]: value,
      },
    }));
  };

  private handleFileChange = (event: React.ChangeEvent<HTMLInputElement>) => {
    const file = event.target.files && event.target.files[0];
    if (file) {
      this.setState({ uploadFile: file });
    }
  };

  private handleChoiceChange = (index: number) => (event: React.ChangeEvent<HTMLInputElement>) => {
    const newChoices = [...this.state.choices];
    newChoices[index] = event.target.value;
    this.setState({ choices: newChoices });
  };

  private addChoiceField = () => {
    this.setState(prevState => ({ choices: [...prevState.choices, ''] }));
  };

  private addMetadataField = () => {
    const newFieldName = this.state.newMetadataField;
    const newFieldType = this.state.newMetadataType;
    const choices = this.state.choices;

    if (this.props.list_title) {
      this._spOperations.AddMetadataField(this.props.context, this.props.list_title, newFieldName, newFieldType, choices)
        .then((result: string) => {
          console.log(result);
          this._spOperations.GetListColumns(this.props.context, this.props.list_title)
            .then((results: SPListColumn[]) => {
              const metadata: { [key: string]: any } = {};
              results.forEach(column => {
                metadata[column.internalName] = '';
              });
              this.setState({
                listColumns: results,
                metadata,
                newMetadataField: '',
                newMetadataDescription: '',
                newMetadataType: 'Text',
                showAddMetadataModal: false,
                choices: [''],
              });
            })
            .catch(error => {
              console.error('Error getting list columns:', error);
            });
        })
        .catch(error => {
          console.error('Error adding metadata field:', error);
        });
    }
  };

  private handleDocumentTypeChange = (event: React.FormEvent<HTMLDivElement>, option?: IDropdownOption) => {
    if (option) {
      this.setState({ selectedDocumentType: option.key as string });
    }
  };

  private filterEmptyFields(fields: { [key: string]: any }): { [key: string]: any } {
    const filteredFields: { [key: string]: any } = {};
    for (const key in fields) {
      if (fields[key] !== null && fields[key] !== undefined && fields[key] !== '') {
        filteredFields[key] = fields[key];
      }
    }
    return filteredFields;
  }
  

  private handleCreateSubmit = () => {
    const fileType = this.state.selectedDocumentType;
    const metadata = this.filterEmptyFields({ ...this.state.metadata });
    const fileName = `${metadata['FileLeafRef']}.${fileType}`; // Utilize the "Nom" metadata
  
    this._spOperations.CreateFile(this.props.context, this.props.list_title, fileName, fileType, metadata)
      .then((result: string) => {
        this.setState({ status: result });
      })
      .catch(error => {
        console.error('Error creating file:', error);
      });
  
    this.setState({ showCreateModal: false, selectedDocumentType: 'txt' });
  };
  
  

  private handleUploadSubmit = () => {
    if (this.state.uploadFile) {
      const metadata = this.filterEmptyFields({ ...this.state.metadata });
      const fileName = this.state.uploadFile.name;
      metadata['FileLeafRef'] = fileName; // Utilize the uploaded file name as "FileLeafRef" metadata
  
      this._spOperations.UploadFile(this.props.context, this.props.list_title, this.state.uploadFile, metadata)
        .then((result: string) => {
          this.setState({ status: result });
        })
        .catch(error => {
          console.error('Error uploading file:', error);
        });
  
      this.setState({ showUploadModal: false, uploadFile: null });
    }
  };
  
  
  private handleCreateFolderSubmit = async () => {
    const {  metadata } = this.state;
    const filteredMetadata = this.filterEmptyFields(metadata);
    const newFolderName = "new folder"
    
      try {
        const result = await this._spOperations.CreateFolder(this.props.context, this.props.list_title, newFolderName, filteredMetadata);
        alert(result);
        this.setState({ showCreateFolderModal: false, newFolderName: '', metadata: {} });
      } catch (error) {
        console.error('Error creating folder:', error);
      }

  };

  private handleFolderNameChange = (event: React.ChangeEvent<HTMLInputElement>) => {
    this.setState({ newFolderName: event.target.value });
  };
  

  private getButtonClass = () => {
    switch (this.props.buttonType) {
      case 'rounded':
        return styles.buttonRounded;
      case 'semi-rounded':
        return styles.buttonSemiRounded;
      case 'strict':
        return styles.buttonStrict;
      default:
        return '';
    }
  };

  public render(): React.ReactElement<IFormulaireDeSoumessionDesDocumentsProps> {
    const { hasTeamsContext, backgroundColor, textColor } = this.props;
  
    const buttonStyle = {
      backgroundColor: backgroundColor,
      color: textColor, // Set the text color for the buttons
    };
  
    const documentTypeOptions: IDropdownOption[] = [
      { key: 'docx', text: 'Word Document (.docx)' },
      { key: 'txt', text: 'Text Document (.txt)' },
      { key: 'pptx', text: 'PowerPoint Presentation (.pptx)' },
      { key: 'xlsx', text: 'Excel Spreadsheet (.xlsx)' },
    ];
    const userOptions: IDropdownOption[] = this.state.users.map(user => ({
      key: user.id,
      text: user.title
    }));
    const columnFields = this.state.listColumns
      .filter(column => !(this.state.isUploadMode && column.internalName === 'FileLeafRef'))
      .map(column => {
        let inputType: string | undefined;
  
        switch (column.type) {
          case 'Text':
            inputType = 'text';
            break;
          case 'Note':
            inputType = 'textarea';
            break;
          case 'Number':
            inputType = 'number';
            break;
          case 'DateTime':
            inputType = 'date';
            break;
          case 'Boolean':
            inputType = 'checkbox';
            break;
          case 'Choice':
            inputType = 'select';
            break;
          case 'User':
            inputType = 'user';
            break;
          case 'URL':
            if (column.displayFormat === 1) {
              inputType = 'file';
            }
            break;
          default:
            inputType = 'text';
        }
  
        if (inputType === 'select') {
          return (
            <div key={column.internalName} className="mb-3">
              {!column.readOnly && (
                <>
                  <label htmlFor={column.internalName}>{column.title}</label>
                  <div className={styles['field-wrapper']}>
                    <div className={styles['field-group']}>
                      <select
                        name={column.internalName}
                        className={styles['select-field']}
                        id={column.internalName}
                        onChange={this.handleSelectChange(column.internalName)}
                      >
                        {column.choices &&
                          column.choices.map(choice => (
                            <option key={choice} value={choice}>
                              {choice}
                            </option>
                          ))}
                      </select>
                    </div>
                  </div>
                </>
              )}
            </div>
          );
        } else if (inputType === 'file') {
          return (
            <div key={column.internalName} className="mb-3">
              {!column.readOnly && (
                <>
                  <label htmlFor={column.internalName}>{column.title}</label>
                  <div className={styles['field-wrapper']}>
                    <div className={styles['field-group']}>
                      <input
                        type="file"
                        id={column.internalName}
                        accept="image/*"
                        className={styles['text-field']}
                        onChange={this.handleFileChange}
                      />
                    </div>
                  </div>
                </>
              )}
            </div>
          );
        } else if (inputType === 'checkbox') {
          return (
            <div key={column.internalName} className="mb-3">
              {!column.readOnly && (
                <div className={styles['field-wrapper']}>
                  <input
                    type="checkbox"
                    id={column.internalName}
                    className={styles['checkbox-field']}
                    onChange={this.handleInputChange(column.internalName)}
                    checked={!!this.state.metadata[column.internalName]}
                  />
                  <label htmlFor={column.internalName}>{column.title}</label>
                </div>
              )}
            </div>
          );
        } else if (inputType === 'textarea') {
          return (
            <div key={column.internalName} className="mb-3">
              {!column.readOnly && (
                <>
                  <label htmlFor={column.internalName}>{column.title}</label>
                  <div className={styles['field-wrapper']}>
                    <div className={styles['field-group']}>
                      <textarea
                        id={column.internalName}
                        className={styles['text-field']}
                        onChange={this.handleTextareaChange(column.internalName)}
                        value={this.state.metadata[column.internalName] || ''}
                      />
                    </div>
                  </div>
                </>
              )}
            </div>
          );
        } else if (inputType === 'user') {
          return (
            <div key={column.internalName} className="mb-3">
              {!column.readOnly && (
                <>
                  <label htmlFor={column.internalName}>{column.title}</label>
                  <div className={styles['field-wrapper']}>
                    <div className={styles['field-group']}>
                      <Dropdown
                        placeholder="Select user"
                        options={userOptions}
                        onChange={(event, option) => this.handleUserChange(event, option, column.internalName)}
                      />
                    </div>
                  </div>
                </>
              )}
            </div>
          );
        } else {
          return (
            <div key={column.internalName} className="mb-3">
              {!column.readOnly && (
                <>
                  <label htmlFor={column.internalName}>{column.title}</label>
                  <div className={styles['field-wrapper']}>
                    <div className={styles['field-group']}>
                      <input
                        type={inputType}
                        id={column.internalName}
                        className={styles['text-field']}
                        onChange={this.handleInputChange(column.internalName)}
                        value={this.state.metadata[column.internalName] || ''}
                      />
                    </div>
                  </div>
                </>
              )}
            </div>
          );
        }
      });
  
    return (
      <section className={`${styles.formulaireDeSoumessionDesDocuments} ${hasTeamsContext ? styles.teams : ''}`}>
        <div className={styles.welcome}>
          <ButtonGrid
            openCreateModal={this.openCreateModal}
            openUploadModal={this.openUploadModal}
            openAddMetadataModal={this.openAddMetadataModal}
            openCreateFolderModal={this.openCreateFolderModal}
            listTitle={this.props.list_title}
            buttonStyle={buttonStyle}
            getButtonClass={this.getButtonClass}
          />
          <CreateDocumentModal
            isOpen={this.state.showCreateModal}
            onDismiss={() => this.setState({ showCreateModal: false })}
            handleCreateSubmit={this.handleCreateSubmit}
            handleDocumentTypeChange={this.handleDocumentTypeChange}
            handleUserChange={this.handleUserChange}
            columnFields={columnFields}
            documentTypeOptions={documentTypeOptions}
            users={this.state.users.map(user => ({ key: user.id, text: user.title }))}
          />
          <UploadDocumentModal
            isOpen={this.state.showUploadModal}
            onDismiss={() => this.setState({ showUploadModal: false })}
            handleUploadSubmit={this.handleUploadSubmit}
            handleFileChange={this.handleFileChange}
            columnFields={columnFields}
          />
          <AddMetadataModal
            isOpen={this.state.showAddMetadataModal}
            onDismiss={() => this.setState({ showAddMetadataModal: false })}
            handleNewMetadataFieldChange={this.handleNewMetadataFieldChange}
            handleNewMetadataDescriptionChange={this.handleNewMetadataDescriptionChange}
            handleNewMetadataTypeChange={this.handleNewMetadataTypeChange}
            handleChoiceChange={this.handleChoiceChange}
            addChoiceField={this.addChoiceField}
            addMetadataField={this.addMetadataField}
            newMetadataField={this.state.newMetadataField}
            newMetadataDescription={this.state.newMetadataDescription}
            newMetadataType={this.state.newMetadataType}
            metadataTypeOptions={metadataTypeOptions}
            choices={this.state.choices}
            choiceGroupOptions={choiceGroupOptions}
          />
          <CreateFolderModal
            isOpen={this.state.showCreateFolderModal}
            onDismiss={() => this.setState({ showCreateFolderModal: false })}
            handleCreateFolderSubmit={this.handleCreateFolderSubmit}
            columnFields={columnFields}
            onFolderNameChange={this.handleFolderNameChange}
          />
        </div>
      </section>
    );
  }
  }


