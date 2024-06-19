import * as React from 'react';
import styles from './BarreDeRecherche.module.scss';
import type { IBarreDeRechercheProps } from './IBarreDeRechercheProps';
import type { IBarreDeRechercheState } from './IBarreDeRechercheState';
import { SPOperations } from '../../Services/SPBarreRechercheService';
import DocumentModal from './DocumentModal';


type Metadata = {
  [key: string]: any;
};

export default class BarreDeRecherche extends React.Component<IBarreDeRechercheProps, IBarreDeRechercheState> {
  private _spOperations: SPOperations;

  constructor(props: IBarreDeRechercheProps) {
    super(props);
    this._spOperations = new SPOperations();
    this.state = {
      listItems: [],
      filteredItems: [],
      status: '',
      searchQuery: '',
      isModalOpen: false,
      selectedDocument: null,
      documentMetadata: null,
      fieldNamesMap: {}, // Initialize the fieldNamesMap
    };
  }

  private handleLibraryChange = (prevSelectedLibraries: string[], newSelectedLibraries: string[]) => {
    // Remove items from deselected libraries
    const removedLibraries = prevSelectedLibraries.filter(lib => !newSelectedLibraries.includes(lib));
    if (removedLibraries.length > 0) {
      const remainingItems = this.state.listItems.filter(item =>
        newSelectedLibraries.includes(item.Library)
      );
      this.setState({ listItems: remainingItems, filteredItems: remainingItems }, () => {
        this.filterItems();
      });
    }

    // Fetch items for newly selected libraries
    const addedLibraries = newSelectedLibraries.filter(lib => !prevSelectedLibraries.includes(lib));
    if (addedLibraries.length > 0) {
      Promise.all(addedLibraries.map(library =>
        this._spOperations.GetListItems(this.props.context, library)
      )).then(results => {
        const mergedResults = results.reduce((acc, val) => acc.concat(val), []);
        this.setState(prevState => ({
          listItems: [...prevState.listItems, ...mergedResults],
          filteredItems: [...prevState.filteredItems, ...mergedResults]
        }), () => {
          this.filterItems();
        });
      }).catch((error: any) => {
        console.error('Error fetching items:', error);
      });
    }
  };

  componentDidUpdate(prevProps: IBarreDeRechercheProps) {
    if (prevProps.selectedLibraries !== this.props.selectedLibraries) {
      this.handleLibraryChange(prevProps.selectedLibraries, this.props.selectedLibraries);
    }
  }

  private handleSearchChange = (event: React.ChangeEvent<HTMLInputElement>) => {
    const searchQuery = event.target.value.toLowerCase();
    this.setState({ searchQuery }, () => {
      this.filterItems();
    });
  };

  private filterItems = () => {
    const { searchQuery, listItems } = this.state;
    if (searchQuery) {
      const filteredItems = listItems.filter(item =>
        (item.Title && item.Title.toLowerCase().includes(searchQuery)) ||
        (item.FileLeafRef && item.FileLeafRef.toLowerCase().includes(searchQuery))
      );
      this.setState({ filteredItems });
    } else {
      this.setState({ filteredItems: listItems });
    }
  };

  private handleSearchSubmit = () => {
    if (this.props.selectedLibraries.length > 0) {
      Promise.all(this.props.selectedLibraries.map(library =>
        this._spOperations.GetListItems(this.props.context, library)
      )).then(results => {
        const mergedResults = results.reduce((acc, val) => acc.concat(val), []); // Flatten the array using reduce
        this.setState({ listItems: mergedResults, filteredItems: mergedResults }, () => {
          this.filterItems();
        });
      }).catch((error: any) => {
        console.error('Error fetching items:', error);
      });
    }
  };

  private handleKeyPress = (event: React.KeyboardEvent<HTMLInputElement>) => {
    if (event.key === 'Enter') {
      this.handleSearchSubmit();
    }
  };

  private fileIconMap: { [key: string]: string } = {
    'xlsx': 'https://res.cdn.office.net/files/fabric-cdn-prod_20240129.001/assets/item-types/20/xlsx.svg',
    'xls': 'https://res.cdn.office.net/files/fabric-cdn-prod_20240129.001/assets/item-types/20/xlsx.svg',
    'doc': 'https://res.cdn.office.net/files/fabric-cdn-prod_20240129.001/assets/item-types/20/docx.svg',
    'docx': 'https://res.cdn.office.net/files/fabric-cdn-prod_20240129.001/assets/item-types/20/docx.svg',
    'ppt': 'https://res.cdn.office.net/files/fabric-cdn-prod_20240129.001/assets/item-types/20/pptx.svg',
    'pptx': 'https://res.cdn.office.net/files/fabric-cdn-prod_20240129.001/assets/item-types/20/pptx.svg',
    'pdf': 'https://res.cdn.office.net/files/fabric-cdn-prod_20240129.001/assets/item-types/20/pdf.svg',
    'txt': 'https://res.cdn.office.net/files/fabric-cdn-prod_20240129.001/assets/item-types/20/txt.svg',
    'jpg': 'https://res.cdn.office.net/files/fabric-cdn-prod_20240129.001/assets/item-types/20/photo.svg',
    'jpeg': 'https://res.cdn.office.net/files/fabric-cdn-prod_20240129.001/assets/item-types/20/photo.svg',
    'png': 'https://res.cdn.office.net/files/fabric-cdn-prod_20240129.001/assets/item-types/20/photo.svg',
    'gif': 'https://res.cdn.office.net/files/fabric-cdn-prod_20240129.001/assets/item-types/20/photo.svg',
    'zip': 'https://res.cdn.office.net/files/fabric-cdn-prod_20240129.001/assets/item-types/20/zip.svg',
    'rar': 'https://res.cdn.office.net/files/fabric-cdn-prod_20240129.001/assets/item-types/20/zip.svg', 
    'default': 'https://res.cdn.office.net/files/fabric-cdn-prod_20240129.001/assets/item-types/20/genericfile.svg' 
  };

  private getFileIconUrl = (fileName: string): string => {
    const extension = fileName.split('.').pop()?.toLowerCase() || '';
    return this.fileIconMap[extension] || this.fileIconMap['default'];
  };

  private openModal = (document: any) => {
    Promise.all([
      this._spOperations.GetListItemMetadata(this.props.context, document.Library, document.ID),
      this._spOperations.GetListFields(this.props.context, document.Library)
    ])
    .then(([metadata, fields]) => {
      const editableFields = fields.filter(field => !field.Hidden && !field.ReadOnlyField);
      const fieldNamesMap = editableFields.reduce((acc, field) => {
        acc[field.InternalName] = field.Title;
        return acc;
      }, {} as { [key: string]: string });
  
      const filteredMetadata: Metadata = Object.keys(metadata).reduce((acc, key) => {
        if (fieldNamesMap[key] && !key.startsWith('@odata')) {
          acc[key] = metadata[key];
        }
        return acc;
      }, {} as Metadata);
  
      const importantFields = ['Modified', 'Editor', 'responsable'];
      importantFields.forEach(field => {
        if (metadata[field] && !fieldNamesMap[field]) {
          fieldNamesMap[field] = field;
          filteredMetadata[field] = metadata[field];
        }
      });
  
      // Ensure FileRef or the correct URL property is added to the document
      document.FileRef = metadata.FileRef;
  
      this.setState({ isModalOpen: true, selectedDocument: document, documentMetadata: filteredMetadata, fieldNamesMap });
    })
    .catch(error => {
      console.error('Error fetching metadata or fields:', error);
      this.setState({ isModalOpen: true, selectedDocument: document, documentMetadata: null, fieldNamesMap: {} });
    });
  };
  
  
  

  private closeModal = () => {
    this.setState({ isModalOpen: false, selectedDocument: null, documentMetadata: null, fieldNamesMap: {} });
  };

  public render(): React.ReactElement<IBarreDeRechercheProps> {
    const { hasTeamsContext, searchLabel, searchLabelColor, searchBarBackground, borderRadiusStyle } = this.props;
    const { filteredItems, searchQuery, isModalOpen, selectedDocument, documentMetadata, fieldNamesMap } = this.state;
    const containerStyle = {
      backgroundColor: searchBarBackground,
      borderRadius: borderRadiusStyle,
      padding: '10px',
    };
    const searchLabelStyle = {
      color: searchLabelColor,
    };
    const searchFieldStyle = {
      borderRadius: borderRadiusStyle,
    };
  
    return (
      <section className={`${styles.ged365Webpart} ${hasTeamsContext ? styles.teams : ''}`} style={containerStyle}>
        <div className={styles.searchWrapper}>
          <label className={styles.searchLabel} style={searchLabelStyle}>{searchLabel}</label>
          <div className={styles.searchContainer} style={searchFieldStyle}>
            <input
              type="text"
              placeholder="Recherche"
              value={searchQuery}
              onChange={this.handleSearchChange}
              onKeyPress={this.handleKeyPress}
              className={styles.searchBox}
            />
            {searchQuery && filteredItems.length > 0 && (
              <div className={styles.dropdown}>
                {filteredItems.map(item => (
                  <div key={item.ID} className={styles.dropdownItem} onClick={() => this.openModal(item)}>
                    {item.FileLeafRef && <img className={styles.icon} src={this.getFileIconUrl(item.FileLeafRef)} alt="" />}
                    {item.FileLeafRef}
                  </div>
                ))}
              </div>
            )}
          </div>
        </div>
        <DocumentModal
          isOpen={isModalOpen}
          onClose={this.closeModal}
          document={selectedDocument}
          documentMetadata={documentMetadata}
          fieldNamesMap={fieldNamesMap}
          getFileIconUrl={this.getFileIconUrl}
        />
      </section>
    );
  }
  
  
}