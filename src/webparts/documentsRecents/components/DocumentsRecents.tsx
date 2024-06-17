import * as React from 'react';
import styles from './DocumentsRecents.module.scss';
import type { IDocumentsRecentsProps } from './IDocumentsRecentsProps';
import type { IDocumentsRecentsState } from './IDocumentsRecentsState';
import { SPOperations } from "../../Services/SPDocRecentsServices";
import { escape } from '@microsoft/sp-lodash-subset';
import DocumentItem from './Documents/DocumentItem';
export default class DocumentsRecents extends React.Component<
  IDocumentsRecentsProps,
  IDocumentsRecentsState
> {
  public _spOperations: SPOperations;

  constructor(props: IDocumentsRecentsProps) {
    super(props);
    this._spOperations = new SPOperations();
    this.state = {
      listTiltes: [],
      listItems: [],
      status: "",
      Titre_list_item: "",
      listItemId: "",
    };
  }

  public componentDidMount() {
    this._fetchFiles();
  }

  public componentDidUpdate(prevProps: IDocumentsRecentsProps) {
    if (prevProps.list_title !== this.props.list_title ||
      prevProps.timeInterval !== this.props.timeInterval ||
      prevProps.numFields !== this.props.numFields) {
      this._fetchFiles();
    }
  }

  private async _fetchFiles(): Promise<void> {
    const { list_title, timeInterval, numFields } = this.props;
    if (list_title) {
      try {
        const files = await this._spOperations.GetRecentFiles(this.props.context, list_title, timeInterval);
        this.setState({ listItems: files.slice(0, numFields) });
      } catch (error) {
        console.error('Error fetching files:', error);
      }
    }
  }

  public render(): React.ReactElement<IDocumentsRecentsProps> {
    const { description, backgroundColor, textColor } = this.props;
    const { listItems } = this.state;

    if (!this.props.list_title) {
      return (
        <>
          <h4>Selectionner la liste que vous souhaiter visualiser</h4>
        </>
      );
    }

    return (
      <section className={`${styles.ged365Webpart}`}>
        <div className={styles.webpartContainer} style={{ backgroundColor }}>
          <div style={{ color: textColor }}>
            <strong>{escape(description)}</strong>
          </div>
          <div>
            {listItems.length > 0 ? (
              <ul className={styles.documentList}>
                {listItems.map((item, index) => (
                  <DocumentItem key={index} item={item} />
                ))}
              </ul>
            ) : (
              <p>No files found for the selected interval.</p>
            )}
          </div>
        </div>
      </section>
    );
  }
}
