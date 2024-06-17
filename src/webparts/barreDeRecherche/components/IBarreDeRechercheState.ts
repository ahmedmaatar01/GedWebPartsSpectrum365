import { SPListItem } from "../../Services/SPBarreRechercheService";

export interface IBarreDeRechercheState {
  listItems: SPListItem[];
  filteredItems: SPListItem[];
  status: string;
  searchQuery: string;
}
