export interface IMemberProfilesProps {
  /** Existing props */
  listId: string;                 // data list (GUID or Title depending on your service)
  imageListId?: string;           // optional image/picture library
  itemsPerPage?: number;          // 0/undefined = show all
  accentColor?: string;           // hex string used for --accent css var

  /** NEW: dynamic headings */
  pageTitle?: string;             // webpart page title
  pageSubtitle?: string;          // webpart page subtitle
}
