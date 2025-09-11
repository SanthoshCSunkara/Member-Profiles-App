import { WebPartContext } from '@microsoft/sp-webpart-base';

export interface IMemberProfilesProps {
  // EXISTING props used throughout your app
  context: WebPartContext;
  listId: string;
  imageListId?: string;
  itemsPerPage?: number;      // 0 or undefined => show all
  accentColor?: string;
  showInMobile?: boolean;     // if you already use this in CSS/logic

  // NEW page header props
  pageTitle: string;
  pageSubtitle: string;
}
