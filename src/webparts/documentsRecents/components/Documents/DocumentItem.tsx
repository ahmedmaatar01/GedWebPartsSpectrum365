import * as React from 'react';
import { Icon } from 'office-ui-fabric-react/lib/Icon';
import styles from './DocumentItem.module.scss';

export interface IDocumentItemProps {
  item: any;
}

const DocumentItem: React.FC<IDocumentItemProps> = ({ item }) => {
  return (
    <li className={styles.documentItem}>
      <div className={styles.icon}>
        <Icon iconName={item.FileSystemObjectType === 1 ? 'FabricFolder' : 'Page'} />
      </div>
      <div className={styles.details}>
        <div><strong>{item.Title || item.FileLeafRef}</strong></div>
        <div>Ajouté le : {new Date(item.Created).toLocaleDateString()}</div>
        <div>Soumis par : {item.Author.Title}</div>
        <div>Modifié par : {item.Editor ? item.Editor.Title : '(si trouvable)'}</div>
      </div>
      <div className={styles.more}>
        <Icon iconName="More" />
      </div>
    </li>
  );
};

export default DocumentItem;
