import { IDropdownOption } from "office-ui-fabric-react";
import { SPListItem, SPListColumn } from "../../Services/SpFormulaireSoumessionService";

export interface IFormulaireDeSoumessionDesDocumentsState {
  listColumns: SPListColumn[];
  listTiltes: IDropdownOption[];
  listItems: SPListItem[];
  status: string;
  Titre_list_item: string;
  showModal: boolean;
  listItemId: string;
  selectedDocumentType: string;
  metadata: { [key: string]: any };
  uploadFile: File | null;
  isUploadMode: boolean;
  showCreateModal: boolean;
  showUploadModal: boolean;
  showAddMetadataModal: boolean;
  newMetadataField: string;
  newMetadataDescription: string;
  newMetadataType: string;
  choices: string[];
  showCreateFolderModal: boolean;
  newFolderName: string;
  users: { id: number, title: string }[]; // Add this line
  selectedUser: number | null; // Add this line
}
