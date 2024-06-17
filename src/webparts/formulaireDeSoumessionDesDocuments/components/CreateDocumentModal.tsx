import * as React from 'react';
import { Modal, Button, Dropdown, IDropdownOption } from 'office-ui-fabric-react';
import styles from './FormulaireDeSoumessionDesDocuments.module.scss';

interface ICreateDocumentModalProps {
  isOpen: boolean;
  onDismiss: () => void;
  handleCreateSubmit: () => void;
  handleDocumentTypeChange: (event: React.FormEvent<HTMLDivElement>, option?: IDropdownOption) => void;
  handleUserChange: (event: React.FormEvent<HTMLDivElement>, option?: IDropdownOption) => void;
  columnFields: JSX.Element[];
  documentTypeOptions: IDropdownOption[];
  users: IDropdownOption[];
}

const CreateDocumentModal: React.FC<ICreateDocumentModalProps> = ({
  isOpen,
  onDismiss,
  handleCreateSubmit,
  handleDocumentTypeChange,
  handleUserChange,
  columnFields,
  documentTypeOptions,
  users
}) => {
  return (
    <Modal
      isOpen={isOpen}
      onDismiss={onDismiss}
      isBlocking={false}
      containerClassName={styles.modalContainer}
    >
      <div className={styles.modalHeader}>
        <span>Créer un document</span>
        <Button iconProps={{ iconName: 'Cancel' }} onClick={onDismiss} />
      </div>
      <div className={styles.modalBody}>
        {columnFields}

        <div className="mb-3">
          <label>Select Document Type</label>
          <div className={styles['field-wrapper']}>
            <div className={styles['field-group']}>
              <Dropdown
                placeholder="Select document type"
                options={documentTypeOptions}
                onChange={handleDocumentTypeChange}
              />
            </div>
          </div>
        </div>

      </div>
      <div className={styles.modalFooter}>
        <Button className={styles.myButton} text="Submit" onClick={handleCreateSubmit} />
        <Button className={styles.cancelButton} text="Cancel" onClick={onDismiss} />
      </div>
    </Modal>
  );
};

export default CreateDocumentModal;
