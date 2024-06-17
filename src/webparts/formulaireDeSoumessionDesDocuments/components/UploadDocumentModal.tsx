import * as React from 'react';
import { Modal, Button } from 'office-ui-fabric-react';
import styles from './FormulaireDeSoumessionDesDocuments.module.scss';

interface IUploadDocumentModalProps {
  isOpen: boolean;
  onDismiss: () => void;
  handleUploadSubmit: () => void;
  handleFileChange: (event: React.ChangeEvent<HTMLInputElement>) => void;
  columnFields: JSX.Element[];
}

const UploadDocumentModal: React.FC<IUploadDocumentModalProps> = ({
  isOpen,
  onDismiss,
  handleUploadSubmit,
  handleFileChange,
  columnFields
}) => {
  return (
    <Modal
      isOpen={isOpen}
      onDismiss={onDismiss}
      isBlocking={false}
      containerClassName={styles.modalContainer}
    >
      <div className={styles.modalHeader}>
        <span>+ Ajouter un document</span>
        <Button iconProps={{ iconName: 'Cancel' }} onClick={onDismiss} />
      </div>
      <div className={styles.modalBody}>
        <div className="mb-3">
          <label htmlFor="uploadFile">Select File to Upload</label>
          <div className={styles['field-wrapper']}>
            <div className={styles['field-group']}>
              <input
                type="file"
                id="uploadFile"
                className={styles['text-field']}
                onChange={handleFileChange}
              />
            </div>
          </div>
        </div>

        {columnFields}
      </div>
      <div className={styles.modalFooter}>
        <Button className={styles.myButton} text="Submit" onClick={handleUploadSubmit} />
        <Button className={styles.cancelButton} text="Cancel" onClick={onDismiss} />
      </div>
    </Modal>
  );
};

export default UploadDocumentModal;
