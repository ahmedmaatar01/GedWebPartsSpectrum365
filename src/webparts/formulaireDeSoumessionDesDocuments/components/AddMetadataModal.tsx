import * as React from 'react';
import { Modal, Button, Dropdown, IDropdownOption, ChoiceGroup, IChoiceGroupOption } from 'office-ui-fabric-react';
import styles from './FormulaireDeSoumessionDesDocuments.module.scss';

interface IAddMetadataModalProps {
  isOpen: boolean;
  onDismiss: () => void;
  handleNewMetadataFieldChange: (event: React.ChangeEvent<HTMLInputElement>) => void;
  handleNewMetadataDescriptionChange: (event: React.ChangeEvent<HTMLInputElement>) => void;
  handleNewMetadataTypeChange: (event: React.FormEvent<HTMLDivElement>, option?: IDropdownOption) => void;
  handleChoiceChange: (index: number) => (event: React.ChangeEvent<HTMLInputElement>) => void;
  addChoiceField: () => void;
  addMetadataField: () => void;
  newMetadataField: string;
  newMetadataDescription: string;
  newMetadataType: string;
  metadataTypeOptions: IDropdownOption[];
  choices: string[];
  choiceGroupOptions: IChoiceGroupOption[];
}

const AddMetadataModal: React.FC<IAddMetadataModalProps> = ({
  isOpen,
  onDismiss,
  handleNewMetadataFieldChange,
  handleNewMetadataDescriptionChange,
  handleNewMetadataTypeChange,
  handleChoiceChange,
  addChoiceField,
  addMetadataField,
  newMetadataField,
  newMetadataDescription,
  newMetadataType,
  metadataTypeOptions,
  choices,
  choiceGroupOptions
}) => {
  return (
    <Modal
      isOpen={isOpen}
      onDismiss={onDismiss}
      isBlocking={false}
      containerClassName={styles.modalContainer}
    >
      <div className={styles.modalHeader}>
        <span>+ Ajouter une métadonnée</span>
        <Button iconProps={{ iconName: 'Cancel' }} onClick={onDismiss} />
      </div>
      <div className={styles.modalBody}>
        <div className="mb-3">
          <label htmlFor="newMetadataField">Name</label>
          <div className={styles['field-wrapper']}>
            <div className={styles['field-group']}>
              <input
                type="text"
                id="newMetadataField"
                className={styles['text-field']}
                onChange={handleNewMetadataFieldChange}
                value={newMetadataField}
              />
            </div>
          </div>
        </div>
        <div className="mb-3">
          <label htmlFor="newMetadataDescription">Description</label>
          <div className={styles['field-wrapper']}>
            <div className={styles['field-group']}>
              <input
                type="text"
                id="newMetadataDescription"
                className={styles['text-field']}
                onChange={handleNewMetadataDescriptionChange}
                value={newMetadataDescription}
              />
            </div>
          </div>
        </div>
        <div className="mb-3">
          <label htmlFor="newMetadataType">Metadata Type</label>
          <div className={styles['field-wrapper']}>
            <div className={styles['field-group']}>
              <Dropdown
                placeholder="Select metadata type"
                options={metadataTypeOptions}
                onChange={handleNewMetadataTypeChange}
                selectedKey={newMetadataType}
              />
            </div>
          </div>
        </div>
        {newMetadataType === 'Text' && (
          <div className="mb-3">
            <label>Text Value</label>
            <div className={styles['field-wrapper']}>
              <div className={styles['field-group']}>
                <input type="text" className={styles['text-field']} value="" readOnly />
              </div>
            </div>
          </div>
        )}
        {newMetadataType === 'Choice' && (
          <>
            {choices.map((choice, index) => (
              <div key={index} className="mb-3">
                <label>Choice {index + 1}</label>
                <div className={styles['field-wrapper']}>
                  <div className={styles['field-group']}>
                    <input
                      type="text"
                      className={styles['text-field']}
                      value={choice}
                      onChange={handleChoiceChange(index)}
                    />
                  </div>
                </div>
              </div>
            ))}
            <Button text="Add Choice" onClick={addChoiceField} />
          </>
        )}
        {newMetadataType === 'Number' && (
          <div className="mb-3">
            <label>Number Value</label>
            <div className={styles['field-wrapper']}>
              <div className={styles['field-group']}>
                <input type="number" className={styles['text-field']} value="" readOnly />
              </div>
            </div>
          </div>
        )}
        {newMetadataType === 'Boolean' && (
          <div className="mb-3">
            <label>Boolean Value</label>
            <div className={styles['field-wrapper']}>
              <div className={styles['field-group']}>
                <ChoiceGroup options={choiceGroupOptions} defaultSelectedKey="No" />
              </div>
            </div>
          </div>
        )}
        {newMetadataType === 'Image' && (
          <div className="mb-3">
            <label>Upload Image</label>
            <div className={styles['field-wrapper']}>
              <div className={styles['field-group']}>
                <input type="file" accept="image/*" className={styles['text-field']} />
              </div>
            </div>
          </div>
        )}
      </div>
      <div className={styles.modalFooter}>
        <Button className={styles.myButton} text="Add" onClick={addMetadataField} />
        <Button className={styles.cancelButton} text="Cancel" onClick={onDismiss} />
      </div>
    </Modal>
  );
};

export default AddMetadataModal;
