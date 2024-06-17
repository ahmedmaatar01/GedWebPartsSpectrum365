import { WebPartContext } from "@microsoft/sp-webpart-base";

export interface IDocumentsRecentsProps {
  description: string;
  isDarkTheme: boolean;
  environmentMessage: string;
  hasTeamsContext: boolean;
  userDisplayName: string;
  context: WebPartContext;
  list_title: string;
  backgroundColor: string;
  textColor: string;
  timeInterval: string; // New property
  numFields: number; // New property
}
