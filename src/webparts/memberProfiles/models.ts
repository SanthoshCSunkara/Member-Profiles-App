export interface IMemberProfilesProps {
  listId: string;
  itemsPerPage: number;
  accentColor: string;
}

export interface IProfileItemRaw {
  Id: number;
  Title: string;
  Role?: string;
  Hire_x0020_Date?: string | Date;
  Birthday?: string;
  CompanyProfile?: any; // FieldUrlValue
  LinkedIn?: any;       // FieldUrlValue
  Image0?: any;         // Image column (internal: Image0)
  About?: string;       // Rich text
  Modified?: string | Date;
  Created?: string | Date;
}

export interface IProfileItem {
  id: number;
  name: string;
  role: string;
  hireDate?: string;
  birthday?: string;
  companyUrl?: string;
  linkedInUrl?: string;
  photoUrl?: string;
  detailsHtml?: string;
}