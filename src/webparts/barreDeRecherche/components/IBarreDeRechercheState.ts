import { SPListItem } from "../../Services/SPBarreRechercheService";


export type Metadata = {
  [key: string]: any;
};

export interface IBarreDeRechercheState {
  listItems: SPListItem[];
  filteredItems: SPListItem[];
  status: string;
  searchQuery: string;
  isModalOpen: boolean;
  selectedDocument: any | null;
  documentMetadata: Metadata | null;
  fieldNamesMap: { [key: string]: string }; 
}
