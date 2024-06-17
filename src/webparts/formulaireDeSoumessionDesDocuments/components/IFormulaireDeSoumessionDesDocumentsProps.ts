import { WebPartContext } from "@microsoft/sp-webpart-base";

export interface IFormulaireDeSoumessionDesDocumentsProps {
  description: string;
  isDarkTheme: boolean;
  environmentMessage: string;
  hasTeamsContext: boolean;
  userDisplayName: string;
  context: WebPartContext;
  list_title: string;
  buttonType: 'rounded' | 'semi-rounded' | 'strict';
  backgroundColor: string;
  textColor: string; // Ensure this property is present
}
