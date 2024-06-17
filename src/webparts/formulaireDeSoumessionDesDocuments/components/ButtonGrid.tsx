import * as React from 'react';
import { Button } from 'office-ui-fabric-react';
import styles from './FormulaireDeSoumessionDesDocuments.module.scss';

interface IButtonGridProps {
  openCreateModal: () => void;
  openUploadModal: () => void;
  openAddMetadataModal: () => void;
  openCreateFolderModal: () => void;
  listTitle: string;
  buttonStyle: React.CSSProperties;
  getButtonClass: () => string;
}

const ButtonGrid: React.FC<IButtonGridProps> = ({
  openCreateModal,
  openUploadModal,
  openAddMetadataModal,
  openCreateFolderModal,
  listTitle,
  buttonStyle,
  getButtonClass
}) => {
  return (
    <div className={styles['button-grid']}>
      <Button
        text="+ Créer un document"
        onClick={openCreateModal}
        className={`${getButtonClass()} ${styles.myButton}`}
        style={buttonStyle}
      />
      <Button
        text="+ Ajouter un document"
        onClick={openUploadModal}
        className={`${getButtonClass()} ${styles.myButton}`}
        style={buttonStyle}
      />
      <Button
        text="+ Ajouter une métadonnée"
        onClick={openAddMetadataModal}
        className={`${getButtonClass()} ${styles.myButton}`}
        style={buttonStyle}
        disabled={!listTitle}
      />
      <Button
        text="+ Créer un dossier"
        onClick={openCreateFolderModal}
        className={`${getButtonClass()} ${styles.myButton}`}
        style={buttonStyle}
        disabled={!listTitle}
      />
    </div>
  );
};

export default ButtonGrid;
