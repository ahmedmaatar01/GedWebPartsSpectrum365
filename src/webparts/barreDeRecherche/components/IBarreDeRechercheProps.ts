import { WebPartContext } from "@microsoft/sp-webpart-base";

export interface IBarreDeRechercheProps {
  context: WebPartContext;
  hasTeamsContext?: boolean;
  isDarkTheme?: boolean;
  searchLabel: string;
  searchLabelColor: string;
  searchBarBackground: string;
  borderRadiusStyle: string;
  selectedLibraries: string[]; 
}
