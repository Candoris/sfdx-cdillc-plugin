import { Connection } from '@salesforce/core';
import { Workbook, Worksheet } from 'exceljs';
import { DescribeGlobalResult, DescribeGlobalSObjectResult, DescribeSObjectResult } from 'jsforce';

interface AppMenuItem {
  Id: string;
  ApplicationId: string;
  Label: string;
  Name: string;
  NamespacePrefix?: string;
}

interface PermissionSet {
  Name: string;
  Label: string;
}

interface Profile {
  Name: string;
}

interface ProfileOrPermissionSetMetadata {
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
  tabSettings: TabSetting | TabSetting[];
  userPermissions: UserPermission | UserPermission[];

  // permission set only
  hasActivationRequired: boolean;
  label: string;
  license: string;

  // profile only
  custom: string;
  categoryGroupVisibilities: ProfileCategoryGroupVisibility | ProfileCategoryGroupVisibility[];
  layoutAssignments: ProfileLayoutAssignment | ProfileLayoutAssignment[];
  userLicense: string;
  loginFlows: ProfileLoginFlow | ProfileLoginFlow[];
  loginHours: ProfileLoginHours | ProfileLoginHours[];
  loginIpRanges: ProfileLoginIpRange | ProfileLoginIpRange[];
  profileActionOverrides: ProfileActionOverride | ProfileActionOverride[];
}

// interface ProfileMetadata {
//   custom: string;
//   description: string;
//   fullName: string;
//   userLicense: string;
//   applicationVisibilities: ApplicationVisibility | ApplicationVisibility[];
//   classAccesses: ApexClassAccess | ApexClassAccess[];
//   customMetadataTypeAccesses: CustomMetadataTypeAccess | CustomMetadataTypeAccess[];
//   customPermissions: CustomPermission | CustomPermission[];
//   customSettingAccesses: CustomSettingAccess | CustomSettingAccess[];
//   externalDataSourceAccesses: ExternalDataSourceAccess | ExternalDataSourceAccess[];
//   fieldPermissions: FieldPermission | FieldPermission[];
//   flowAccesses: FlowAccess | FlowAccess[];
//   layoutAssignments: ProfileLayoutAssignment | ProfileLayoutAssignment[];
//   objectPermissions: ObjectPermission | ObjectPermission[];
//   pageAccesses: ApexPageAccess | ApexPageAccess[];
//   recordTypeVisibilities: RecordTypeVisibility | RecordTypeVisibility[];
//   tabSettings: TabSetting | TabSetting[];
//   userPermissions: UserPermission | UserPermission[];
// }

interface ApplicationVisibility {
  application: string;
  visible: string;
  default: string; // only profiles
}

interface ApexClassAccess {
  apexClass: string;
  enabled: string;
}

interface CustomMetadataTypeAccess {
  enabled: string;
  name: string;
}

interface CustomPermission {
  enabled: string;
  name: string;
}

interface CustomSettingAccess {
  enabled: string;
  name: string;
}

interface ExternalDataSourceAccess {
  enabled: string;
  externalDataSource: string;
}

interface FieldPermission {
  editable: string;
  field: string;
  readable: string;
  hidden: string; // only profiles
}

interface FlowAccess {
  enabled: string;
  flow: string;
}

interface ProfileLayoutAssignment {
  layout: string;
  recordType: string;
}

interface ObjectPermission {
  allowCreate: string;
  allowDelete: string;
  allowEdit: string;
  allowRead: string;
  modifyAllRecords: string;
  object: string;
  viewAllRecords: string;
}

interface ApexPageAccess {
  apexPage: string;
  enabled: string;
}

interface RecordTypeVisibility {
  recordType: string;
  visible: string;
}

interface TabSetting {
  tab: string;
  visibility: string;
}

interface UserPermission {
  enabled: string;
  name: string;
}

interface ProfileLoginFlow {
  flow: string;
  flowtype: string;
  friendlyname: string;
  uiLoginFlowType: string;
  useLightningRuntime: string;
  vfFlowPage: string;
  vfFlowPageTitle: string;
}

interface ProfileActionOverride {
  actionName: string;
  content: string;
  formFactor: string;
  pageOrSobjectType: string;
  recordType: string;
  type: string;
}

interface ProfileApplicationVisibility {
  application: string;
  default: string;
  visible: string;
}

interface ProfileCategoryGroupVisibility {
  dataCategories: string[];
  dataCategorGroup: string;
  visibility: string;
}

interface ProfileLoginHours {
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

interface ProfileLoginIpRange {
  description: string;
  endAddress: string;
  startAddress: string;
}

export default class PermissionsExportBuilder {
  private conn: Connection;
  private describeGlobal: DescribeGlobalResult;
  private sObjectDescribeMap: Map<string, DescribeSObjectResult>;
  private appMenuItemsMap: Map<string, AppMenuItem>;

  public constructor(conn: Connection) {
    this.conn = conn;
    this.sObjectDescribeMap = new Map<string, DescribeSObjectResult>();
  }

  public generatePermissionsXLS = async (
    permissionSetNames: string[],
    profileNames: string[],
    filePath: string
  ): Promise<void> => {
    const workbook = new Workbook();
    const [describeGlobal, psObjDescribe, appMenuItemsMap] = await Promise.all([
      this.conn.describeGlobal(),
      this.conn.describe('PermissionSet'),
      this.queryAppMenuItemsMap(),
    ]);
    this.describeGlobal = describeGlobal;
    this.sObjectDescribeMap.set('PermissionSet', psObjDescribe);
    this.appMenuItemsMap = appMenuItemsMap;

    const promises = [];
    if (permissionSetNames?.length) {
      promises.push(this.createPermissionSetSheets(permissionSetNames, workbook));
    }

    if (profileNames?.length) {
      promises.push(this.createProfileSheets(profileNames, workbook));
    }

    if (promises.length) {
      await Promise.all(promises);
      await workbook.xlsx.writeFile(filePath);
    } else {
      throw new Error('Permission file not created. No permission set names or profile names provided.');
    }
  };

  private getGlobalSObjectDescribeByName = (name: string): DescribeGlobalSObjectResult => {
    return this.describeGlobal.sobjects.find((s) => {
      return s.name === name;
    });
  };

  private buildSheetHeader = (sheet: Worksheet, startCellKey: string, endCellKey: string, title: string): void => {
    sheet.mergeCells(`${startCellKey}:${endCellKey}`);
    sheet.getCell(startCellKey).value = title;
    sheet.getCell(startCellKey).font = {
      bold: true,
      size: 24,
    };
    sheet.getCell(startCellKey).alignment = {
      vertical: 'middle',
      horizontal: 'center',
    };
  };

  private addHeaderRow = (sheet: Worksheet, colVals: string[]): void => {
    const row = sheet.addRow(colVals);
    row.font = {
      bold: true,
      size: 16,
    };
  };

  private addSubheaderRow = (sheet: Worksheet, colVals: string[]): void => {
    const row = sheet.addRow(colVals);
    row.font = {
      bold: true,
      size: 14,
    };
  };

  private addDetailRow = (sheet: Worksheet, colVals: string[]): void => {
    const row = sheet.addRow(colVals);
    row.font = {
      size: 12,
    };
  };

  private addApplicationVisibilities = (
    sheet: Worksheet,
    applicationVisibilities: ApplicationVisibility[],
    appMenuItemMap: Map<string, AppMenuItem>
  ): void => {
    if (applicationVisibilities?.length) {
      this.addHeaderRow(sheet, ['Assigned Apps']);
      this.addSubheaderRow(sheet, ['Label', 'API Name', 'Default']);

      applicationVisibilities.forEach((av) => {
        const appMenuItem = appMenuItemMap.get(av.application);
        if (appMenuItem) {
          this.addDetailRow(sheet, [appMenuItem.Label, appMenuItem.Name, av.default]);
        }
      });
      sheet.addRow(['']);
    }
  };

  private addApexClassAccesses = (sheet: Worksheet, apexClassAccesses: ApexClassAccess[]): void => {
    if (apexClassAccesses?.length) {
      this.addHeaderRow(sheet, ['Apex Class Accesses']);
      this.addSubheaderRow(sheet, ['Apex Class', 'Enabled']);

      apexClassAccesses.sort((a, b) => {
        return a.apexClass.localeCompare(b.apexClass);
      });
      apexClassAccesses.forEach((aca) => {
        if (aca.enabled === 'true') {
          this.addDetailRow(sheet, [aca.apexClass, aca.enabled]);
        }
      });
      sheet.addRow(['']);
    }
  };

  private addCustomMetadataTypeAccesses = (
    sheet: Worksheet,
    customMetadataTypeAccesses: CustomMetadataTypeAccess[]
  ): void => {
    if (customMetadataTypeAccesses?.length) {
      this.addHeaderRow(sheet, ['Custom Metadata Type Accesses']);
      this.addSubheaderRow(sheet, ['Name', 'Enabled']);

      customMetadataTypeAccesses.sort((a, b) => {
        return a.name.localeCompare(b.name);
      });
      customMetadataTypeAccesses.forEach((cmt) => {
        if (cmt.enabled === 'true') {
          this.addDetailRow(sheet, [cmt.name, cmt.enabled]);
        }
      });
      sheet.addRow(['']);
    }
  };

  private addCustomPermissions = (sheet: Worksheet, customPermissions: CustomPermission[]): void => {
    if (customPermissions?.length) {
      this.addHeaderRow(sheet, ['Custom Permissions']);
      this.addSubheaderRow(sheet, ['Name', 'Enabled']);

      customPermissions.sort((a, b) => {
        return a.name.localeCompare(b.name);
      });

      customPermissions.forEach((cp) => {
        if (cp.enabled === 'true') {
          this.addDetailRow(sheet, [cp.name, cp.enabled]);
        }
      });
      sheet.addRow(['']);
    }
  };

  private addCustomSettingAccesses = (sheet: Worksheet, customSettingAccesses: CustomSettingAccess[]): void => {
    if (customSettingAccesses?.length) {
      this.addHeaderRow(sheet, ['Custom Setting Accesses']);
      this.addSubheaderRow(sheet, ['Name', 'Enabled']);

      customSettingAccesses.sort((a, b) => {
        return a.name.localeCompare(b.name);
      });
      customSettingAccesses.forEach((csa) => {
        if (csa.enabled === 'true') {
          this.addDetailRow(sheet, [csa.name, csa.enabled]);
        }
      });
      sheet.addRow(['']);
    }
  };

  private addFlowAccesses = (sheet: Worksheet, flowAccesses: FlowAccess[]): void => {
    if (flowAccesses?.length) {
      this.addHeaderRow(sheet, ['Flow Accesses']);
      this.addSubheaderRow(sheet, ['Name', 'Enabled']);

      flowAccesses.sort((a, b) => {
        return a.flow.localeCompare(b.flow);
      });
      flowAccesses.forEach((fa) => {
        if (fa.enabled === 'true') {
          this.addDetailRow(sheet, [fa.flow, fa.enabled]);
        }
      });
      sheet.addRow(['']);
    }
  };

  private addPageLayoutAssignments = (
    sheet: Worksheet,
    pageLayoutAssignments: ProfileLayoutAssignment[],
    objectPermissions: ObjectPermission[]
  ): void => {
    if (pageLayoutAssignments?.length) {
      this.addHeaderRow(sheet, ['Page Layout Assignments']);
      this.addSubheaderRow(sheet, ['Object Label', 'Object API Name', 'Record Type', 'Page Layout Assignment']);

      const plaRows = [];
      pageLayoutAssignments.forEach((pla) => {
        const [layoutObjectName, layoutName] = pla.layout.split('-');
        const objPerms = objectPermissions.find((op) => {
          return op.object === layoutObjectName;
        });

        if (objPerms) {
          const sobj = this.getGlobalSObjectDescribeByName(layoutObjectName);

          let recordType = 'Master';
          if (pla.recordType) {
            const [, rtName] = pla.recordType.split('.');
            recordType = rtName;
          }

          if (sobj && objPerms) {
            plaRows.push([sobj.label, sobj.name, recordType, layoutName]);
          }
        }
      });

      // sort by label, record type
      plaRows.sort((a: string[], b: string[]) => {
        return a[0].localeCompare(b[0]) || a[2].localeCompare(b[2]);
      });

      plaRows.forEach((plaRow) => {
        this.addDetailRow(sheet, plaRow);
      });
      sheet.addRow(['']);
    }
  };

  private addObjectPermissions = (sheet: Worksheet, objectPermissions: ObjectPermission[]): void => {
    if (objectPermissions?.length) {
      this.addHeaderRow(sheet, ['Object Permissions']);
      this.addSubheaderRow(sheet, ['Label', 'API Name', 'Permission']);

      const opRows = [];
      objectPermissions.forEach((op) => {
        const permissions = [];
        if (op.allowRead === 'true') {
          permissions.push('Read');
        }
        if (op.allowCreate === 'true') {
          permissions.push('Create');
        }
        if (op.allowEdit === 'true') {
          permissions.push('Edit');
        }
        if (op.allowDelete === 'true') {
          permissions.push('Delete');
        }
        if (op.viewAllRecords === 'true') {
          permissions.push('View All');
        }
        if (op.modifyAllRecords === 'true') {
          permissions.push('Modify All');
        }
        const sobj = this.getGlobalSObjectDescribeByName(op.object);
        opRows.push([sobj?.label, op?.object, permissions.join('/')]);
      });

      opRows.sort((a: string[], b: string[]) => {
        return a[0].localeCompare(b[0]);
      });

      opRows.forEach((opRow) => {
        this.addDetailRow(sheet, opRow);
      });
      sheet.addRow(['']);
    }
  };

  private addFieldPermissions = async (
    sheet: Worksheet,
    fieldPermissions: FieldPermission[],
    objectPermissions: ObjectPermission[]
  ): Promise<void> => {
    if (fieldPermissions?.length) {
      this.addHeaderRow(sheet, ['Field Level Permissions']);
      this.addSubheaderRow(sheet, ['Object Label', 'Object API Name', 'Field Label', 'Field API Name', 'Permission']);

      const fieldRows = [];
      await Promise.all(
        fieldPermissions.map((fp) => {
          const [objAPIName, fieldAPIName] = fp.field.split('.');
          const objPerms = objectPermissions.find((op) => op.object === objAPIName);

          // only show fields where the profile has at least read access
          if (objPerms) {
            let permission: string;
            if (fp.readable === 'true') {
              permission = 'Read';
            }
            if (fp.editable === 'true') {
              permission = 'Edit';
            }
            const sobj = this.getGlobalSObjectDescribeByName(objAPIName);
            if (permission) {
              fieldRows.push([
                sobj.label,
                sobj.name,
                '', // fieldDescribe?.label,
                fieldAPIName,
                permission,
              ]);
            }
          }
        })
      );
      fieldRows.sort((a: string[], b: string[]) => {
        return a[0].localeCompare(b[0]) || a[3].localeCompare(b[3]);
      });
      fieldRows.forEach((fieldRow) => {
        this.addDetailRow(sheet, fieldRow);
      });
      sheet.addRow(['']);
    }
  };

  private addPageAccesses = (sheet: Worksheet, pageAccesses: ApexPageAccess[]): void => {
    if (pageAccesses?.length) {
      this.addHeaderRow(sheet, ['Visualforce Page Accesses']);
      this.addSubheaderRow(sheet, ['Name', 'Enabled']);

      pageAccesses.sort((a, b) => {
        return a.apexPage.localeCompare(b.apexPage);
      });
      pageAccesses.forEach((apa) => {
        if (apa.enabled === 'true') {
          this.addDetailRow(sheet, [apa.apexPage, apa.enabled]);
        }
      });
      sheet.addRow(['']);
    }
  };

  private addRecordTypeVisibilities = (sheet: Worksheet, recordTypeVisibilities: RecordTypeVisibility[]): void => {
    if (recordTypeVisibilities?.length) {
      this.addHeaderRow(sheet, ['Record Type Visibilities']);
      this.addSubheaderRow(sheet, ['Name', 'Visibility']);

      recordTypeVisibilities.sort((a, b) => {
        return a.recordType.localeCompare(b.recordType);
      });
      recordTypeVisibilities.forEach((rtv) => {
        this.addDetailRow(sheet, [rtv.recordType, rtv.visible]);
      });
      sheet.addRow(['']);
    }
  };

  private addTabVisibilities = (sheet: Worksheet, tabVisibilities: TabSetting[]): void => {
    if (tabVisibilities?.length) {
      this.addHeaderRow(sheet, ['Tab Visibilities']);
      this.addSubheaderRow(sheet, ['Name', 'Visibility']);

      tabVisibilities.sort((a, b) => {
        return a.tab.localeCompare(b.tab);
      });
      tabVisibilities.forEach((tv) => {
        this.addDetailRow(sheet, [tv.tab, tv.visibility]);
      });
      sheet.addRow(['']);
    }
  };

  private addUserPermissions = (sheet: Worksheet, userPermissions: UserPermission[]): void => {
    if (userPermissions?.length) {
      this.addHeaderRow(sheet, ['System Permissions']);
      this.addSubheaderRow(sheet, ['Permission', 'Access']);

      userPermissions.sort((a, b) => {
        return a.name.localeCompare(b.name);
      });
      userPermissions.forEach((up) => {
        this.addDetailRow(sheet, [up.name, 'true']);
      });
      sheet.addRow(['']);
    }
  };

  private buildWhereInStringValue = (arr: string[]): string => {
    if (!arr || !arr.length) {
      return '';
    }
    return arr
      .map((val) => {
        return `'${val}'`;
      })
      .join(',');
  };

  private queryAppMenuItemsMap = async (): Promise<Map<string, AppMenuItem>> => {
    const soql = 'SELECT Id, ApplicationId, Label, Name, NamespacePrefix FROM AppMenuItem ORDER BY Label ASC';
    return this.conn.query(soql).then((result) => {
      const appMenuItemMap: Map<string, AppMenuItem> = new Map<string, AppMenuItem>();
      result.records.forEach((rec: AppMenuItem) => {
        appMenuItemMap.set(rec.Name, rec);
      });
      return appMenuItemMap;
    });
  };

  private queryPermissionSets = async (permissionSetNames: string[]): Promise<PermissionSet[]> => {
    const psNamesStr = this.buildWhereInStringValue(permissionSetNames);
    const soql = `SELECT Name, Label FROM PermissionSet WHERE Name IN (${psNamesStr}) ORDER BY Label ASC`;
    return this.conn.query(soql).then((result) => result.records as PermissionSet[]);
  };

  private queryProfiles = async (profileNames: string[]): Promise<Profile[]> => {
    const psNamesStr = this.buildWhereInStringValue(profileNames);
    const soql = `SELECT Name FROM Profile WHERE Name IN (${psNamesStr}) ORDER BY Name ASC`;
    return this.conn.query(soql).then((result) => result.records as Profile[]);
  };

  private getMetadataPropAsArray = <T>(prop: string, metadata: unknown): T[] => {
    if (Array.isArray(metadata[prop])) {
      return metadata[prop] as T[];
    } else if (!Array.isArray(metadata[prop]) && typeof metadata[prop] === 'object' && metadata[prop] !== null) {
      return [metadata[prop]] as T[];
    } else {
      return metadata[prop] as T[];
    }
  };

  private addPermissionsToSheet = async (
    sheet: Worksheet,
    metadataRecord: ProfileOrPermissionSetMetadata,
    isProfile: boolean
  ): Promise<void> => {
    this.addApplicationVisibilities(
      sheet,
      this.getMetadataPropAsArray('applicationVisibilities', metadataRecord),
      this.appMenuItemsMap
    );
    this.addApexClassAccesses(sheet, this.getMetadataPropAsArray('classAccesses', metadataRecord));
    this.addCustomMetadataTypeAccesses(
      sheet,
      this.getMetadataPropAsArray('customMetadataTypeAccesses', metadataRecord)
    );
    this.addCustomPermissions(sheet, this.getMetadataPropAsArray('customPermissions', metadataRecord));
    this.addCustomSettingAccesses(sheet, this.getMetadataPropAsArray('customSettingAccesses', metadataRecord));
    this.addFlowAccesses(sheet, this.getMetadataPropAsArray('flowAccesses', metadataRecord));

    if (metadataRecord.objectPermissions) {
      if (isProfile) {
        this.addPageLayoutAssignments(
          sheet,
          this.getMetadataPropAsArray('layoutAssignments', metadataRecord),
          this.getMetadataPropAsArray('objectPermissions', metadataRecord)
        );
      }

      this.addObjectPermissions(sheet, this.getMetadataPropAsArray('objectPermissions', metadataRecord));

      await this.addFieldPermissions(
        sheet,
        this.getMetadataPropAsArray('fieldPermissions', metadataRecord),
        this.getMetadataPropAsArray('objectPermissions', metadataRecord)
      );
    }

    this.addPageAccesses(sheet, this.getMetadataPropAsArray('pageAccesses', metadataRecord));
    this.addRecordTypeVisibilities(sheet, this.getMetadataPropAsArray('recordTypeVisibilities', metadataRecord));
    this.addTabVisibilities(sheet, this.getMetadataPropAsArray('tabSettings', metadataRecord));
    this.addUserPermissions(sheet, this.getMetadataPropAsArray('userPermissions', metadataRecord));
  };

  private getProfileOrPermissionSetData = async (
    metadataType: string,
    fullNames: string[]
  ): Promise<ProfileOrPermissionSetMetadata[]> => {
    const metadataResult = (await this.conn.metadata.read(metadataType, fullNames)) as unknown;
    let metadataRecords: ProfileOrPermissionSetMetadata[];
    if (Array.isArray(metadataResult)) {
      metadataRecords = metadataResult as ProfileOrPermissionSetMetadata[];
    } else {
      metadataRecords = [metadataResult as ProfileOrPermissionSetMetadata];
    }
    return metadataRecords;
  };

  private createPermissionSetSheets = async (permissionSetNames: string[], workbook: Workbook): Promise<void> => {
    const permissionSets: PermissionSet[] = await this.queryPermissionSets(permissionSetNames);
    if (!permissionSets?.length) {
      return;
    }
    const validPermissionSetNames = permissionSets.map((ps) => ps.Name);
    const metadataRecords = await this.getProfileOrPermissionSetData('PermissionSet', validPermissionSetNames);

    await Promise.all(
      permissionSets.map(async (ps) => {
        const metadataRecord = metadataRecords.find((psm) => {
          return psm.fullName === ps.Name;
        });
        const sheet = workbook.addWorksheet(ps['Label']);

        for (let i = 1; i <= 5; i++) {
          sheet.getColumn(i).width = 50;
        }

        this.buildSheetHeader(sheet, 'A1', 'C1', `Permission Set: ${ps['Label']}`);
        sheet.addRow(['']);

        await this.addPermissionsToSheet(sheet, metadataRecord, false);
      })
    );
  };

  private createProfileSheets = async (profileNames: string[], workbook: Workbook): Promise<void> => {
    const profiles: Profile[] = await this.queryProfiles(profileNames);
    if (!profiles?.length) {
      return;
    }
    const validProfileNames = profiles.map((ps) => ps.Name);
    const metadataRecords = await this.getProfileOrPermissionSetData('Profile', validProfileNames);

    await Promise.all(
      profiles.map(async (p) => {
        const profile = profiles.find((prof) => prof.Name === p.Name);
        const metadataRecord = metadataRecords.find((mr) => mr.fullName === profile.Name);
        const sheet = workbook.addWorksheet(profile.Name);

        for (let i = 1; i <= 5; i++) {
          sheet.getColumn(i).width = 50;
        }

        this.buildSheetHeader(sheet, 'A1', 'C1', `Profile: ${profile.Name}`);
        sheet.addRow(['']);

        await this.addPermissionsToSheet(sheet, metadataRecord, true);
      })
    );
  };
}
