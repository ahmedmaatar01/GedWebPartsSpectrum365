import React, { useState, useEffect } from 'react';
import { Modal, Button, Form, Spinner, Alert } from 'react-bootstrap';
import { SPListItem, SPListColumn, SPOperations } from "../../../Services/SpTableauDeBoardService";

interface IEditModalProps {
    show: boolean;
    handleClose: () => void;
    item: SPListItem | null;
    columns: SPListColumn[];
    context: any;
    listTitle: string;
    handleSave: (updatedItem: SPListItem) => void;
}

const EditModal: React.FC<IEditModalProps> = ({ show, handleClose, item, columns, context, listTitle, handleSave }) => {
    const [formData, setFormData] = useState<SPListItem>({} as SPListItem);
    const [fileNameWithoutExtension, setFileNameWithoutExtension] = useState<string>('');
    const [fileExtension, setFileExtension] = useState<string>('');
    const [users, setUsers] = useState<{ id: number, title: string }[]>([]);
    const [error, setError] = useState<string | null>(null);
    const [loading, setLoading] = useState<boolean>(false);
    const [successMessage, setSuccessMessage] = useState<string | null>(null);

    useEffect(() => {
        if (item) {
            setFormData(item);
            const fileName = item.FileLeafRef;
            const lastDotIndex = fileName.lastIndexOf('.');
            if (lastDotIndex !== -1) {
                setFileNameWithoutExtension(fileName.substring(0, lastDotIndex));
                setFileExtension(fileName.substring(lastDotIndex));
            } else {
                setFileNameWithoutExtension(fileName);
                setFileExtension('');
            }
        }
    }, [item]);

    useEffect(() => {
        if (columns.some(column => column.type === 'User')) {
            fetchUsers();
        }
    }, [columns]);

    const fetchUsers = async () => {
        try {
            const users = await new SPOperations().GetUsers(context);
            setUsers(users);
        } catch (error) {
            setError('Error fetching users.');
            console.error(error);
        }
    };

    const handleChange = (e: React.ChangeEvent<HTMLInputElement>) => {
        const { name, value } = e.target;
        if (name === 'FileLeafRef') {
            if (!value.includes('.')) {
                setFileNameWithoutExtension(value);
            }
        } else if (name.endsWith("Id")) {
            const fieldName = name.replace("Id", "");
            setFormData(prevState => ({
                ...prevState,
                [fieldName]: value,
                [name]: value // Ensure the state is updated correctly
            }));
        } else {
            setFormData(prevState => ({
                ...prevState,
                [name]: value
            }));
        }
    };

    const handleSubmit = async () => {
        if (!item) return;

        setError(null);
        setSuccessMessage(null);
        setLoading(true);

        const fieldsToUpdate: { [key: string]: any } = {};

        for (const key in formData) {
            if (formData[key] !== item[key] && key !== '@odata.type' && key !== '@odata.id' && key !== '@odata.etag' && key !== '@odata.editLink') {
                if (key === 'FileLeafRef') {
                    fieldsToUpdate[key] = `${fileNameWithoutExtension}${fileExtension}`;
                } else if (columns.some(column => column.internalName === key && column.type === 'User')) {
                    fieldsToUpdate[`${key}Id`] = formData[key];
                } else {
                    fieldsToUpdate[key] = formData[key];
                }
            }
        }

        console.log("Updating item with fields:", fieldsToUpdate);

        try {
            await new SPOperations().UpdateListItemFields(context, listTitle, item.Id, fieldsToUpdate);
            handleSave({ ...item, ...fieldsToUpdate });
            setSuccessMessage('Item updated successfully.');
            handleClose();
        } catch (error) {
            setError('Error updating item.');
            console.error(error);
        } finally {
            setLoading(false);
        }
    };

    return (
        <Modal show={show} onHide={handleClose}>
            <Modal.Header closeButton>
                <Modal.Title>Edit File Information</Modal.Title>
            </Modal.Header>
            <Modal.Body>
                {error && <Alert variant="danger">{error}</Alert>}
                {successMessage && <Alert variant="success">{successMessage}</Alert>}
                {item && (
                    <Form>
                        {columns.map((heading, idx) => (
                            <Form.Group key={idx} className="mb-3">
                                <Form.Label>{heading.title}</Form.Label>
                                {heading.internalName === 'FileLeafRef' ? (
                                    <Form.Control
                                        type="text"
                                        name="FileLeafRef"
                                        value={fileNameWithoutExtension}
                                        onChange={handleChange}
                                    />
                                ) : heading.type === "Choice" ? (
                                    <Form.Control
                                        as="select"
                                        name={heading.internalName}
                                        value={formData[heading.internalName] || ''}
                                        onChange={handleChange}
                                    >
                                        <option value="">Select...</option>
                                        {heading.choices && heading.choices.map(choice => (
                                            <option key={choice} value={choice}>{choice}</option>
                                        ))}
                                    </Form.Control>
                                ) : heading.type === "User" ? (
                                    <Form.Control
                                        as="select"
                                        name={`${heading.internalName}Id`}
                                        value={formData[`${heading.internalName}Id`] || ''}
                                        onChange={handleChange}
                                    >
                                        <option value="">Select a user...</option>
                                        {users.map(user => (
                                            <option key={user.id} value={user.id}>{user.title}</option>
                                        ))}
                                    </Form.Control>
                                ) : (
                                    <Form.Control
                                        type="text"
                                        name={heading.internalName}
                                        value={formData[heading.internalName] || ''}
                                        onChange={handleChange}
                                    />
                                )}
                            </Form.Group>
                        ))}
                    </Form>
                )}
            </Modal.Body>
            <Modal.Footer>
                <Button variant="secondary" onClick={handleClose} disabled={loading}>
                    Close
                </Button>
                <Button variant="primary" onClick={handleSubmit} disabled={loading}>
                    {loading ? <Spinner as="span" animation="border" size="sm" role="status" aria-hidden="true" /> : 'Save Changes'}
                </Button>
            </Modal.Footer>
        </Modal>
    );
};

export default EditModal;
