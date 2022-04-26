import { Connection } from '@salesforce/core';
import { Workbook, Worksheet } from 'exceljs';
import { DescribeGlobalResult, DescribeGlobalSObjectResult, DescribeSObjectResult } from 'jsforce';
import {
  AppMenuItem,
  PermissionSet,
  Profile,
  ProfileOrPermissionSetMetadata,
  ApplicationVisibility,
  ApexClassAccess,
  CustomMetadataTypeAccess,
  CustomPermission,
  CustomSettingAccess,
  FieldPermission,
  FlowAccess,
  ProfileLayoutAssignment,
  ObjectPermission,
  ApexPageAccess,
  RecordTypeVisibility,
  TabSetting,
  UserPermission,
} from './permissionsModels';

export default class PermissionsExportBuilder {
  private conn: Connection;
  private includedComponents: string[];
  private describeGlobal: DescribeGlobalResult;
  private sObjectDescribeMap: Map<string, DescribeSObjectResult>;
  private appMenuItemsMap: Map<string, AppMenuItem>;

  public constructor(conn: Connection, includedComponents: string[]) {
    this.conn = conn;
    this.includedComponents = includedComponents || [];
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

  private isComponentIncluded = (componentName: string): boolean => {
    return (
      !this.includedComponents.length ||
      this.includedComponents.indexOf('all') !== -1 ||
      this.includedComponents.indexOf(componentName) !== -1
    );
  };

  private getGlobalSObjectDescribeByName = (name: string): DescribeGlobalSObjectResult => {
    return this.describeGlobal.sobjects.find((s) => {
      return s.name === name;
    });
  };

  private buildSheetHeader = (sheet: Worksheet, title: string): void => {
    sheet.mergeCells('A1:C1');
    sheet.getCell('A1').value = title;
    sheet.getCell('A1').font = {
      bold: true,
      size: 24,
    };
    sheet.getCell('A1').alignment = {
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

      const rows = [];
      applicationVisibilities.forEach((av) => {
        const appMenuItem = appMenuItemMap.get(av.application);
        if (appMenuItem && av.visible === 'true') {
          let name = appMenuItem.Name;
          if (appMenuItem.NamespacePrefix) {
            name = `${appMenuItem.NamespacePrefix}__${name}`;
          }
          rows.push([appMenuItem.Label, name, av.default]);
        }
      });

      rows.sort((a: string[], b: string[]) => {
        return a[0].localeCompare(b[0]);
      });

      rows.forEach((row) => {
        this.addDetailRow(sheet, row);
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
      const rows = [];

      flowAccesses.sort((a, b) => {
        return a.flow.localeCompare(b.flow);
      });
      flowAccesses.forEach((fa) => {
        if (fa.enabled === 'true') {
          rows.push([fa.flow, fa.enabled]);
        }
      });

      if (rows.length) {
        this.addHeaderRow(sheet, ['Flow Accesses']);
        this.addSubheaderRow(sheet, ['Name', 'Enabled']);

        rows.forEach((row) => {
          this.addDetailRow(sheet, row);
        });

        sheet.addRow(['']);
      }
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

  private addRecordTypeVisibilities = (
    sheet: Worksheet,
    recordTypeVisibilities: RecordTypeVisibility[],
    isProfile: boolean
  ): void => {
    if (recordTypeVisibilities?.length) {
      this.addHeaderRow(sheet, ['Record Type Visibilities']);

      if (isProfile) {
        this.addSubheaderRow(sheet, ['Name', 'Visible', 'Default']);
      } else {
        this.addSubheaderRow(sheet, ['Name', 'Visible']);
      }

      recordTypeVisibilities.sort((a, b) => {
        return a.recordType.localeCompare(b.recordType);
      });
      recordTypeVisibilities.forEach((rtv) => {
        const row = [rtv.recordType, rtv.visible];
        if (isProfile) {
          row.push(rtv.default);
        }
        this.addDetailRow(sheet, row);
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
        let name = rec.Name;
        if (rec.NamespacePrefix) {
          name = `${rec.NamespacePrefix}__${name}`;
        }
        appMenuItemMap.set(name, rec);
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
    if (this.isComponentIncluded('applicationVisibilities')) {
      this.addApplicationVisibilities(
        sheet,
        this.getMetadataPropAsArray('applicationVisibilities', metadataRecord),
        this.appMenuItemsMap
      );
    }
    if (this.isComponentIncluded('classAccesses')) {
      this.addApexClassAccesses(sheet, this.getMetadataPropAsArray('classAccesses', metadataRecord));
    }
    if (this.isComponentIncluded('customMetadataTypeAccesses')) {
      this.addCustomMetadataTypeAccesses(
        sheet,
        this.getMetadataPropAsArray('customMetadataTypeAccesses', metadataRecord)
      );
    }
    if (this.isComponentIncluded('customPermissions')) {
      this.addCustomPermissions(sheet, this.getMetadataPropAsArray('customPermissions', metadataRecord));
    }
    if (this.isComponentIncluded('customSettingAccesses')) {
      this.addCustomSettingAccesses(sheet, this.getMetadataPropAsArray('customSettingAccesses', metadataRecord));
    }
    if (this.isComponentIncluded('flowAccesses')) {
      this.addFlowAccesses(sheet, this.getMetadataPropAsArray('flowAccesses', metadataRecord));
    }

    if (metadataRecord.objectPermissions) {
      if (isProfile && this.isComponentIncluded('layoutAssignments')) {
        this.addPageLayoutAssignments(
          sheet,
          this.getMetadataPropAsArray('layoutAssignments', metadataRecord),
          this.getMetadataPropAsArray('objectPermissions', metadataRecord)
        );
      }

      if (this.isComponentIncluded('objectPermissions')) {
        this.addObjectPermissions(sheet, this.getMetadataPropAsArray('objectPermissions', metadataRecord));
      }

      if (this.isComponentIncluded('fieldPermissions')) {
        await this.addFieldPermissions(
          sheet,
          this.getMetadataPropAsArray('fieldPermissions', metadataRecord),
          this.getMetadataPropAsArray('objectPermissions', metadataRecord)
        );
      }
    }

    if (this.isComponentIncluded('pageAccesses')) {
      this.addPageAccesses(sheet, this.getMetadataPropAsArray('pageAccesses', metadataRecord));
    }
    if (this.isComponentIncluded('recordTypeVisibilities')) {
      this.addRecordTypeVisibilities(
        sheet,
        this.getMetadataPropAsArray('recordTypeVisibilities', metadataRecord),
        isProfile
      );
    }
    if (this.isComponentIncluded('tabSettings')) {
      this.addTabVisibilities(sheet, this.getMetadataPropAsArray('tabSettings', metadataRecord));
    }
    if (this.isComponentIncluded('userPermissions')) {
      this.addUserPermissions(sheet, this.getMetadataPropAsArray('userPermissions', metadataRecord));
    }
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

  private createPermissionsSheet = (workbook: Workbook, sheetName: string, sheetTitle: string): Worksheet => {
    const sheet = workbook.addWorksheet(sheetName);

    for (let i = 1; i <= 5; i++) {
      sheet.getColumn(i).width = 50;
    }

    this.buildSheetHeader(sheet, sheetTitle);
    sheet.addRow(['']);

    return sheet;
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
        const metadataRecord = metadataRecords.find((psm) => psm.fullName === ps.Name);
        const sheet = this.createPermissionsSheet(workbook, ps.Label, `Permission Set: ${ps.Label}`);
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
      profiles.map(async (profile) => {
        const metadataRecord = metadataRecords.find((mr) => mr.fullName === profile.Name);
        const sheet = this.createPermissionsSheet(workbook, profile.Name, `Profile: ${profile.Name}`);
        await this.addPermissionsToSheet(sheet, metadataRecord, true);
      })
    );
  };
}
