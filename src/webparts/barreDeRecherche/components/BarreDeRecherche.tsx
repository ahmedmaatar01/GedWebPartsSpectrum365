import * as React from 'react';
import styles from './BarreDeRecherche.module.scss';
import type { IBarreDeRechercheProps } from './IBarreDeRechercheProps';
import type { IBarreDeRechercheState } from './IBarreDeRechercheState';
import { SPOperations } from '../../Services/SPBarreRechercheService';

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
    };
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
        Object.values(item).some(value =>
          value && value.toString().toLowerCase().includes(searchQuery)
        )
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
  public render(): React.ReactElement<IBarreDeRechercheProps> {
    const { hasTeamsContext, searchLabel, searchLabelColor, searchBarBackground, borderRadiusStyle } = this.props;
    const { filteredItems, searchQuery } = this.state;
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
                  <div key={item.ID} className={styles.dropdownItem}>
                    {item.Title} ({item.FileLeafRef})
                  </div>
                ))}
              </div>
            )}
          </div>
        </div>
      </section>
    );
  }
  
}
