import * as React from 'react';
import { Modal, Button } from 'office-ui-fabric-react';
import styles from './FormulaireDeSoumessionDesDocuments.module.scss';

interface ICreateFolderModalProps {
  isOpen: boolean;
  onDismiss: () => void;
  handleCreateFolderSubmit: () => void;
  columnFields: JSX.Element[];
  onFolderNameChange: (event: React.ChangeEvent<HTMLInputElement>) => void;
}

const CreateFolderModal: React.FC<ICreateFolderModalProps> = ({
  isOpen,
  onDismiss,
  handleCreateFolderSubmit,
  columnFields,
  onFolderNameChange
}) => {
  return (
    <Modal
      isOpen={isOpen}
      onDismiss={onDismiss}
      isBlocking={false}
      containerClassName={styles.modalContainer}
    >
      <div className={styles.modalHeader}>
        <span>Créer un dossier</span>
        <Button iconProps={{ iconName: 'Cancel' }} onClick={onDismiss} />
      </div>
      <div className={styles.modalBody}>

        {columnFields}
      </div>
      <div className={styles.modalFooter}>
        <Button className={styles.myButton} text="Submit" onClick={handleCreateFolderSubmit} />
        <Button className={styles.cancelButton} text="Cancel" onClick={onDismiss} />
      </div>
    </Modal>
  );
};

export default CreateFolderModal;
