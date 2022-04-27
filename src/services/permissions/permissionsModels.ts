export interface AppMenuItem {
  Id: string;
  ApplicationId: string;
  IsActive: string;
  Label: string;
  Name: string;
  NamespacePrefix?: string;
}

export interface RecordType {
  Id: string;
  DeveloperName: string;
  Name: string;
  NamespacePrefix: string;
  SobjectType: string;
}

export interface PermissionSet {
  Name: string;
  Label: string;
}

export interface Profile {
  Name: string;
}

export interface ProfileOrPermissionSetMetadata {
  // shared properties
  description: string;
  fullName: string;
  applicationVisibilities: ApplicationVisibility | ApplicationVisibility[];
  classAccesses: ApexClassAccess | ApexClassAccess[];
  customMetadataTypeAccesses: CustomMetadataTypeAccess | CustomMetadataTypeAccess[];
  customPermissions: CustomPermission | CustomPermission[];
  customSettingAccesses: CustomSettingAccess | CustomSettingAccess[];
  externalDataSourceAccesses: ExternalDataSourceAccess | ExternalDataSourceAccess[];
  fieldPermissions: FieldPermission | FieldPermission[];
  flowAccesses: FlowAccess | FlowAccess[];
  objectPermissions: ObjectPermission | ObjectPermission[];
  pageAccesses: ApexPageAccess | ApexPageAccess[];
  recordTypeVisibilities: RecordTypeVisibility | RecordTypeVisibility[];
  userPermissions: UserPermission | UserPermission[];

  // permission set only
  hasActivationRequired: boolean;
  label: string;
  license: string;
  tabSettings: TabSetting | TabSetting[];

  // profile only
  custom: string;
  categoryGroupVisibilities: ProfileCategoryGroupVisibility | ProfileCategoryGroupVisibility[];
  layoutAssignments: ProfileLayoutAssignment | ProfileLayoutAssignment[];
  userLicense: string;
  loginFlows: ProfileLoginFlow | ProfileLoginFlow[];
  loginHours: ProfileLoginHours | ProfileLoginHours[];
  loginIpRanges: ProfileLoginIpRange | ProfileLoginIpRange[];
  tabVisibilities: TabSetting | TabSetting[];
}

export interface ApplicationVisibility {
  application: string;
  visible: string;
  default: string; // only profiles
}

export interface ApexClassAccess {
  apexClass: string;
  enabled: string;
}

export interface CustomMetadataTypeAccess {
  enabled: string;
  name: string;
}

export interface CustomPermission {
  enabled: string;
  name: string;
}

export interface CustomSettingAccess {
  enabled: string;
  name: string;
}

export interface ExternalDataSourceAccess {
  enabled: string;
  externalDataSource: string;
}

export interface FieldPermission {
  editable: string;
  field: string;
  readable: string;
  hidden: string; // only profiles
}

export interface FlowAccess {
  enabled: string;
  flow: string;
}

export interface ProfileLayoutAssignment {
  layout: string;
  recordType: string;
}

export interface ObjectPermission {
  allowCreate: string;
  allowDelete: string;
  allowEdit: string;
  allowRead: string;
  modifyAllRecords: string;
  object: string;
  viewAllRecords: string;
}

export interface ApexPageAccess {
  apexPage: string;
  enabled: string;
}

export interface RecordTypeVisibility {
  recordType: string;
  visible: string;
  default: string; // profile only
  personAccountDefault: string; // profile only
}

export interface TabSetting {
  tab: string;
  visibility: string;
}

export interface UserPermission {
  enabled: string;
  name: string;
}

export interface ProfileLoginFlow {
  flow: string;
  flowtype: string;
  friendlyname: string;
  uiLoginFlowType: string;
  useLightningRuntime: string;
  vfFlowPage: string;
  vfFlowPageTitle: string;
}

export interface ProfileCategoryGroupVisibility {
  dataCategories: string[];
  dataCategorGroup: string;
  visibility: string;
}

export interface ProfileLoginHours {
  sundayStart: string;
  mondayStart: string;
  tuesdayStart: string;
  wednesdayStart: string;
  thursdayStart: string;
  fridayStart: string;
  saturdayStart: string;
  sundayEnd: string;
  mondayEnd: string;
  tuesdayEnd: string;
  wednesdayEnd: string;
  thursdayEnd: string;
  fridayEnd: string;
  saturdayEnd: string;
}

export interface ProfileLoginIpRange {
  description: string;
  endAddress: string;
  startAddress: string;
}
