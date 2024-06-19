import * as React from 'react';
import { Modal, IconButton } from 'office-ui-fabric-react';
import styles from './DocumentModal.module.scss';

export type Metadata = {
  [key: string]: any;
};

interface IDocumentModalProps {
  isOpen: boolean;
  onClose: () => void;
  document: any;
  documentMetadata: Metadata | null;
  fieldNamesMap: { [key: string]: string };
  getFileIconUrl: (fileName: string) => string;
}

const DocumentModal: React.FC<IDocumentModalProps> = ({ isOpen, onClose, document, documentMetadata, fieldNamesMap, getFileIconUrl }) => {
  if (!document) {
    return null;
  }

  return (
    <Modal
      isOpen={isOpen}
      onDismiss={onClose}
      isBlocking={false}
      containerClassName={styles.modalContainer}
    >
      <div className={styles.modalHeader}>
        <span>{document.FileLeafRef}</span>
        <IconButton
          iconProps={{ iconName: 'Cancel' }}
          ariaLabel="Close"
          onClick={onClose}
        />
      </div>
      <div className={styles.modalBody}>
        <img src={getFileIconUrl(document.FileLeafRef)} alt="" className={styles.icon} />
        <h2>{document.FileLeafRef}</h2>
        {document.FileRef && (
          <div>
            <a href={document.FileRef} target="_blank" rel="noopener noreferrer" className={styles.accessLink}>
              Accéder au document
            </a>
          </div>
        )}
        <div className={styles.metadata}>
          {documentMetadata ? (
            Object.keys(documentMetadata).map((key) => (
              <p key={key}>
                <strong>{fieldNamesMap[key] || key}:</strong> {documentMetadata[key]}
              </p>
            ))
          ) : (
            <p>No metadata available.</p>
          )}
        </div>
      </div>
    </Modal>
  );
};

export default DocumentModal;
