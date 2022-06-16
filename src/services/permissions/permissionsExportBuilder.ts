import { UX } from '@salesforce/command';
import { Connection } from '@salesforce/core';
import { Workbook, Worksheet } from 'exceljs';
import { DescribeGlobalResult, DescribeGlobalSObjectResult, DescribeSObjectResult } from 'jsforce';
import { buildWhereInStringValue, getMetadataPropAsArray, getMetadataAsArray } from '../../utils';
import {
  AppMenuItem,
  RecordType,
  PermissionSet,
  PermissionSetAssignment,
  Profile,
  PermissionSetGroup,
  PermissionSetGroupMetadata,
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
  User,
} from './permissionsModels';
import PermissionSetGroupCombine from './permissionSetGroupCombine';
import UserPermissionsCombine from './userPermissionsCombine';

export default class PermissionsExportBuilder {
  private conn: Connection;
  private ux: UX;
  private includedComponents: string[];
  private describeGlobal: DescribeGlobalResult;
  private sObjectDescribeMap: Map<string, DescribeSObjectResult>;
  private appMenuItemsMap: Map<string, AppMenuItem>;
  private recordTypes: RecordType[];
  private verbose: boolean;

  public constructor(conn: Connection, ux: UX, includedComponents: string[], verbose: boolean) {
    this.conn = conn;
    this.ux = ux;
    this.includedComponents = includedComponents || [];
    this.verbose = verbose;
    this.sObjectDescribeMap = new Map<string, DescribeSObjectResult>();
  }

  public log = (data: string): void => {
    if (this.verbose) {
      this.ux.log(data);
    }
  };

  public generatePermissionsXLS = async (
    permissionSetNames: string[],
    profileNames: string[],
    permissionSetGroupNames: string[],
    userNames: string[],
    filePath: string
  ): Promise<void> => {
    this.log('Preparing spreadsheets and initial data');
    const workbook = new Workbook();
    const [describeGlobal, psObjDescribe, appMenuItemsMap, recordTypes] = await Promise.all([
      this.conn.describeGlobal(),
      this.conn.describe('PermissionSet'),
      this.queryAppMenuItemsMap(),
      this.queryRecordTypes(),
    ]);
    this.describeGlobal = describeGlobal;
    this.sObjectDescribeMap.set('PermissionSet', psObjDescribe);
    this.appMenuItemsMap = appMenuItemsMap;
    this.recordTypes = recordTypes;

    const promises = [];
    if (permissionSetNames?.length) {
      const permissionSets: PermissionSet[] = await this.queryPermissionSets(permissionSetNames);
      if (permissionSets?.length) {
        promises.push(this.createPermissionSetSheets(permissionSets, workbook));
      }
    }

    if (profileNames?.length) {
      const profiles: Profile[] = await this.queryProfiles(profileNames);
      if (profiles?.length) {
        promises.push(this.createProfileSheets(profiles, workbook));
      }
    }

    if (permissionSetGroupNames?.length) {
      const permissionSetGroups: PermissionSetGroup[] = await this.queryPermissionSetGroups(permissionSetGroupNames);
      if (permissionSetGroups?.length) {
        promises.push(this.createPermissionSetGroupSheets(permissionSetGroups, workbook));
      }
    }

    if (userNames?.length) {
      const users: User[] = await this.queryUsers(userNames);
      if (users?.length) {
        promises.push(this.createUserSheets(users, workbook));
      }
    }

    if (promises.length) {
      await Promise.all(promises);
      await workbook.xlsx.writeFile(filePath);
    } else {
      throw new Error(
        'Permissions file not created. No valid profile, permission set, or permission set group names provided.'
      );
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

  private addApplicationVisibilities = (sheet: Worksheet, applicationVisibilities: ApplicationVisibility[]): void => {
    if (applicationVisibilities?.length) {
      const rows = [];
      applicationVisibilities.forEach((av) => {
        const appMenuItem = this.appMenuItemsMap.get(av.application);
        if (appMenuItem && av.visible === 'true') {
          let name = appMenuItem.Name;
          if (appMenuItem.NamespacePrefix) {
            name = `${appMenuItem.NamespacePrefix}__${name}`;
          }
          rows.push([appMenuItem.Label || '', name, av.default || 'false']);
        }
      });

      if (rows.length) {
        this.log(`Adding application visibilities for ${sheet.name}`);
        this.addHeaderRow(sheet, ['Assigned Apps']);
        this.addSubheaderRow(sheet, ['Label', 'API Name', 'Default']);

        rows.sort((a: string[], b: string[]) => {
          return a[0].localeCompare(b[0]);
        });

        rows.forEach((row) => {
          this.addDetailRow(sheet, row);
        });

        sheet.addRow(['']);
        this.log(`Finished adding ${rows.length} application visibilities for ${sheet.name}`);
      }
    }
  };

  private addApexClassAccesses = (sheet: Worksheet, apexClassAccesses: ApexClassAccess[]): void => {
    if (apexClassAccesses?.length) {
      const rows = [];
      apexClassAccesses.forEach((aca) => {
        if (aca.enabled === 'true') {
          rows.push([aca.apexClass || '', aca.enabled || '']);
        }
      });

      if (rows.length) {
        this.log(`Adding apex class accesses for ${sheet.name}`);
        this.addHeaderRow(sheet, ['Apex Class Access']);
        this.addSubheaderRow(sheet, ['Apex Class', 'Enabled']);

        rows.sort((a: string[], b: string[]) => {
          return a[0].localeCompare(b[0]);
        });
        rows.forEach((row) => {
          this.addDetailRow(sheet, row);
        });

        sheet.addRow(['']);
        this.log(`Finished adding ${rows.length} apex class accesses for ${sheet.name}`);
      }
    }
  };

  private addCustomMetadataTypeAccesses = (
    sheet: Worksheet,
    customMetadataTypeAccesses: CustomMetadataTypeAccess[]
  ): void => {
    if (customMetadataTypeAccesses?.length) {
      const rows = [];
      customMetadataTypeAccesses.forEach((cmt) => {
        if (cmt.enabled === 'true') {
          rows.push([cmt.name || '', cmt.enabled || '']);
        }
      });

      if (rows.length) {
        this.log(`Adding custom metadata type accesses for ${sheet.name}`);
        this.addHeaderRow(sheet, ['Custom Metadata Type Accesses']);
        this.addSubheaderRow(sheet, ['Name', 'Enabled']);

        rows.sort((a: string[], b: string[]) => {
          return a[0].localeCompare(b[0]);
        });
        rows.forEach((row) => {
          this.addDetailRow(sheet, row);
        });

        sheet.addRow(['']);
        this.log(`Finished adding ${rows.length} custom metadata type accesses for ${sheet.name}`);
      }
    }
  };

  private addCustomPermissions = (sheet: Worksheet, customPermissions: CustomPermission[]): void => {
    if (customPermissions?.length) {
      const rows = [];
      customPermissions.forEach((cp) => {
        if (cp.enabled === 'true') {
          rows.push([cp.name || '', cp.enabled || '']);
        }
      });

      if (rows.length) {
        this.log(`Adding custom permissions for ${sheet.name}`);
        this.addHeaderRow(sheet, ['Custom Permissions']);
        this.addSubheaderRow(sheet, ['Name', 'Enabled']);

        rows.sort((a: string[], b: string[]) => {
          return a[0].localeCompare(b[0]);
        });
        rows.forEach((row) => {
          this.addDetailRow(sheet, row);
        });

        sheet.addRow(['']);
        this.log(`Finished adding ${rows.length} custom permissions for ${sheet.name}`);
      }
    }
  };

  private addCustomSettingAccesses = (sheet: Worksheet, customSettingAccesses: CustomSettingAccess[]): void => {
    if (customSettingAccesses?.length) {
      const rows = [];
      customSettingAccesses.forEach((csa) => {
        if (csa.enabled === 'true') {
          rows.push([csa.name || '', csa.enabled || '']);
        }
      });

      if (rows.length) {
        this.log(`Adding custom setting accesses for ${sheet.name}`);
        this.addHeaderRow(sheet, ['Custom Setting Accesses']);
        this.addSubheaderRow(sheet, ['Name', 'Enabled']);

        rows.sort((a: string[], b: string[]) => {
          return a[0].localeCompare(b[0]);
        });
        rows.forEach((row) => {
          this.addDetailRow(sheet, row);
        });

        sheet.addRow(['']);
        this.log(`Finished adding ${rows.length} custom setting accesses for ${sheet.name}`);
      }
    }
  };

  private addFlowAccesses = (sheet: Worksheet, flowAccesses: FlowAccess[]): void => {
    if (flowAccesses?.length) {
      const rows = [];
      flowAccesses.forEach((fa) => {
        if (fa.enabled === 'true') {
          rows.push([fa.flow || '', fa.enabled || '']);
        }
      });

      if (rows.length) {
        this.log(`Adding flow accesses for ${sheet.name}`);
        this.addHeaderRow(sheet, ['Flow Accesses']);
        this.addSubheaderRow(sheet, ['Name', 'Enabled']);

        rows.sort((a: string[], b: string[]) => {
          return a[0].localeCompare(b[0]);
        });
        rows.forEach((row) => {
          this.addDetailRow(sheet, row);
        });

        sheet.addRow(['']);
        this.log(`Finished adding ${rows.length} flow accesses for ${sheet.name}`);
      }
    }
  };

  private addPageLayoutAssignments = (sheet: Worksheet, pageLayoutAssignments: ProfileLayoutAssignment[]): void => {
    if (pageLayoutAssignments?.length) {
      const rows = [];
      pageLayoutAssignments.forEach((pla) => {
        const [layoutObjectName, ...layoutNameParts] = pla.layout.split('-');
        const sobj = this.getGlobalSObjectDescribeByName(layoutObjectName);

        let recordType: RecordType;
        if (pla.recordType) {
          const [, rtName] = pla.recordType.split('.');
          recordType = this.recordTypes.find((rt) => {
            return rt.DeveloperName === rtName && rt.SobjectType === sobj.name;
          });
        }

        if (sobj) {
          rows.push([
            sobj?.label || '',
            sobj?.name || '',
            recordType?.Name || 'Master',
            recordType?.DeveloperName || 'Master',
            layoutNameParts.join('-') || '',
          ]);
        }
      });

      if (rows.length) {
        this.log(`Adding page layout accesses for ${sheet.name}`);
        this.addHeaderRow(sheet, ['Page Layout Assignments']);
        this.addSubheaderRow(sheet, [
          'Object Label',
          'Object API Name',
          'Record Type Label',
          'Record Type Developer Name',
          'Page Layout Assignment',
        ]);

        // sort by label, record type
        rows.sort((a: string[], b: string[]) => {
          return a[0].localeCompare(b[0]) || a[2].localeCompare(b[2]);
        });
        rows.forEach((row) => {
          this.addDetailRow(sheet, row);
        });

        sheet.addRow(['']);
        this.log(`Finished adding ${rows.length} page layout accesses for ${sheet.name}`);
      }
    }
  };

  private addObjectPermissions = (sheet: Worksheet, objectPermissions: ObjectPermission[]): void => {
    if (objectPermissions?.length) {
      const rows = [];
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
        rows.push([sobj?.label || '', op?.object || '', permissions.join('/') || '']);
      });

      if (rows.length) {
        this.log(`Adding object permissions for ${sheet.name}`);
        this.addHeaderRow(sheet, ['Object Permissions']);
        this.addSubheaderRow(sheet, ['Label', 'API Name', 'Permission']);

        rows.sort((a: string[], b: string[]) => {
          return a[0].localeCompare(b[0]);
        });
        rows.forEach((opRow) => {
          this.addDetailRow(sheet, opRow);
        });

        sheet.addRow(['']);
        this.log(`Finished adding ${rows.length} object permissions for ${sheet.name}`);
      }
    }
  };

  private buildFieldPermissionRow = (fp: FieldPermission): string[] => {
    const [objAPIName, fieldAPIName] = fp.field.split('.');
    let permission: string;
    if (fp.readable === 'true') {
      permission = 'Read';
    }
    if (fp.editable === 'true') {
      permission = 'Edit';
    }
    // let sObjDescribe = this.sObjectDescribeMap.get(objAPIName);
    // if (!sObjDescribe) {
    //   sObjDescribe = await this.conn.describe(objAPIName);
    //   this.sObjectDescribeMap.set(sObjDescribe.name, sObjDescribe);
    // }
    if (permission) {
      // const fieldDescribe = sObjDescribe.fields.find((f) => f.name === fieldAPIName);
      return [
        // sObjDescribe?.label || '',
        objAPIName || '',
        // fieldDescribe?.label || '',
        fieldAPIName || '',
        permission || '',
      ] as string[];
    }
  };

  private addFieldPermissions = (
    sheet: Worksheet,
    fieldPermissions: FieldPermission[],
    objectPermissions?: ObjectPermission[]
  ): void => {
    if (fieldPermissions?.length) {
      const rows = [];

      fieldPermissions.map((fp) => {
        const [objAPIName] = fp.field.split('.');
        let addRow = false;
        if (objectPermissions?.length) {
          const objPerms = objectPermissions.find((op) => op.object === objAPIName);
          addRow = !!objPerms;
        } else {
          addRow = true;
        }
        if (addRow) {
          const row = this.buildFieldPermissionRow(fp);
          if (row) {
            rows.push(row);
          }
        }
      });

      if (rows.length) {
        this.log(`Adding field permissions for ${sheet.name}`);
        this.addHeaderRow(sheet, ['Field Level Permissions']);
        this.addSubheaderRow(sheet, ['Object API Name', 'Field API Name', 'Permission']);

        rows.sort((a: string[], b: string[]) => {
          return a[0].localeCompare(b[0]) || a[1].localeCompare(b[1]);
        });
        rows.forEach((row) => {
          this.addDetailRow(sheet, row);
        });

        sheet.addRow(['']);
        this.log(`Finished adding ${rows.length} field permissions for ${sheet.name}`);
      }
    }
  };

  private addPageAccesses = (sheet: Worksheet, pageAccesses: ApexPageAccess[]): void => {
    if (pageAccesses?.length) {
      const rows = [];
      pageAccesses.forEach((apa) => {
        if (apa.enabled === 'true') {
          rows.push([apa.apexPage || '', apa.enabled || '']);
        }
      });

      if (rows.length) {
        this.log(`Adding visualforce page accesses for ${sheet.name}`);
        this.addHeaderRow(sheet, ['Visualforce Page Access']);
        this.addSubheaderRow(sheet, ['Name', 'Enabled']);

        rows.sort((a: string[], b: string[]) => {
          return a[0].localeCompare(b[0]);
        });
        rows.forEach((row) => {
          this.addDetailRow(sheet, row);
        });

        sheet.addRow(['']);
        this.log(`Finished adding ${rows.length} visualforce page accesses for ${sheet.name}`);
      }
    }
  };

  private addRecordTypeVisibilities = (
    sheet: Worksheet,
    recordTypeVisibilities: RecordTypeVisibility[],
    isProfile: boolean
  ): void => {
    if (recordTypeVisibilities?.length) {
      const rows = [];
      recordTypeVisibilities.forEach((rtv) => {
        const [objectAPIName, rtDevName] = rtv.recordType.split('.');
        const recordType = this.recordTypes.find((rt) => {
          return rt.DeveloperName === rtDevName && rt.SobjectType === objectAPIName;
        });
        const sObjectDescribe = this.describeGlobal.sobjects.find((s) => {
          return s.name === objectAPIName;
        });

        if (rtv.visible === 'true') {
          const row = [
            sObjectDescribe?.label || '',
            objectAPIName || '',
            recordType?.Name || '',
            recordType?.DeveloperName || rtDevName || '',
            rtv.visible || '',
          ];
          if (isProfile) {
            row.push(rtv.default);
          }
          rows.push(row);
        }
      });

      if (rows.length) {
        this.log(`Adding record type visibilities for ${sheet.name}`);
        this.addHeaderRow(sheet, ['Record Type Visibilities']);
        const subheaderRow = [
          'Object Label',
          'Object API Name',
          'Record Type Label',
          'Record Type Developer Name',
          'Visible',
        ];

        if (isProfile) {
          subheaderRow.push('Default');
        }
        this.addSubheaderRow(sheet, subheaderRow);

        rows.sort((a: string[], b: string[]) => {
          return a[0].localeCompare(b[0]) || a[2].localeCompare(b[2]);
        });
        rows.forEach((row) => {
          this.addDetailRow(sheet, row);
        });

        sheet.addRow(['']);
        this.log(`Finished adding ${rows.length} record type visibilities for ${sheet.name}`);
      }
    }
  };

  private addTabVisibilities = (sheet: Worksheet, tabVisibilities: TabSetting[]): void => {
    if (tabVisibilities?.length) {
      const rows = [];
      tabVisibilities.forEach((tv) => {
        if (tv.visibility !== 'None' && tv.visibility !== 'Hidden') {
          rows.push([tv.tab || '', tv.visibility || '']);
        }
      });

      if (rows.length) {
        this.log(`Adding tab visibilities for ${sheet.name}`);
        this.addHeaderRow(sheet, ['Tab Visibilities']);
        this.addSubheaderRow(sheet, ['Name', 'Visibility']);

        rows.sort((a: string[], b: string[]) => {
          return a[0].localeCompare(b[0]);
        });
        rows.forEach((row) => {
          this.addDetailRow(sheet, row);
        });

        sheet.addRow(['']);
        this.log(`Finished adding ${rows.length} tab visibilities for ${sheet.name}`);
      }
    }
  };

  private addUserPermissions = (sheet: Worksheet, userPermissions: UserPermission[]): void => {
    if (userPermissions?.length) {
      const rows = [];
      userPermissions.forEach((up) => {
        rows.push([up.name || '', 'true']);
      });

      if (rows.length) {
        this.log(`Adding user permissions for ${sheet.name}`);
        this.addHeaderRow(sheet, ['User Permissions']);
        this.addSubheaderRow(sheet, ['Permission', 'Access']);

        rows.sort((a: string[], b: string[]) => {
          return a[0].localeCompare(b[0]);
        });
        rows.forEach((row) => {
          this.addDetailRow(sheet, row);
        });

        sheet.addRow(['']);
        this.log(`Finished adding ${rows.length} user permissions for ${sheet.name}`);
      }
    }
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

  private queryRecordTypes = async (): Promise<RecordType[]> => {
    const soql = 'SELECT Id, DeveloperName, IsActive, Name, NamespacePrefix, SobjectType FROM RecordType';
    return this.conn.query(soql).then((result) => {
      return result.records as RecordType[];
    });
  };

  private queryPermissionSets = async (permissionSetNames: string[]): Promise<PermissionSet[]> => {
    const soql = 'SELECT NamespacePrefix, Name, Label, Type FROM PermissionSet';
    const permissionSets: PermissionSet[] = await this.conn
      .query(soql)
      .then((result) => result.records as PermissionSet[]);

    const filteredPermissionSets: PermissionSet[] = [];
    permissionSetNames.forEach((psName) => {
      let [namespacePrefix, name] = psName.split('__');
      if (!name) {
        name = namespacePrefix;
        namespacePrefix = undefined;
      }

      const permissionSet = permissionSets.find((ps) => {
        // these are tied to profiles and permission set groups
        if (ps.Type === 'Profile' || ps.Type === 'Group') {
          return;
        }
        if (namespacePrefix) {
          return ps.NamespacePrefix === namespacePrefix && ps.Name === name;
        } else {
          return ps.Name === name;
        }
      });
      if (permissionSet) {
        filteredPermissionSets.push(permissionSet);
      }
    });

    filteredPermissionSets.sort((a: PermissionSet, b: PermissionSet) => {
      return a.Label.localeCompare(b.Label);
    });

    return filteredPermissionSets;
  };

  private queryProfiles = async (profileNames: string[]): Promise<Profile[]> => {
    const profileNamesStr = buildWhereInStringValue(profileNames);
    const soql = `SELECT Name FROM Profile WHERE Name IN (${profileNamesStr}) ORDER BY Name ASC`;
    return this.conn.query(soql).then((result) => result.records as Profile[]);
  };

  private queryPermissionSetGroups = async (psgNames: string[]): Promise<PermissionSetGroup[]> => {
    const psNamesStr = buildWhereInStringValue(psgNames);
    const soql = `SELECT DeveloperName, MasterLabel FROM PermissionSetGroup WHERE DeveloperName IN (${psNamesStr}) ORDER BY MasterLabel ASC`;
    return this.conn.query(soql).then((result) => result.records as PermissionSetGroup[]);
  };

  private queryUsers = async (userNames: string[]): Promise<User[]> => {
    const userNamesStr = buildWhereInStringValue(userNames);
    const soql = `SELECT Id, Name, ProfileId, Profile.Name, Username FROM User WHERE Username IN (${userNamesStr}) ORDER BY Username ASC`;
    return this.conn.query(soql).then((result) => result.records as User[]);
  };

  private queryPermissionSetAssignments = async (userIds: string[]): Promise<PermissionSetAssignment[]> => {
    const userIdsStr = buildWhereInStringValue(userIds);
    const soql = `SELECT
    Id,
    AssigneeId,
    IsActive,
    PermissionSetGroupId,
    PermissionSetGroup.DeveloperName,
    PermissionSetGroup.MasterLabel,
    PermissionSetId,
    PermissionSet.Name,
    PermissionSet.Label,
    PermissionSet.Type
    FROM PermissionSetAssignment
    WHERE AssigneeId IN (${userIdsStr}) AND IsActive = TRUE`;
    return this.conn.query(soql).then((result) => result.records as PermissionSetAssignment[]);
  };

  private addPermissionsToSheet = (
    sheet: Worksheet,
    metadataRecord: ProfileOrPermissionSetMetadata,
    isProfile: boolean
  ): void => {
    if (this.isComponentIncluded('applicationVisibilities')) {
      this.addApplicationVisibilities(sheet, getMetadataPropAsArray('applicationVisibilities', metadataRecord));
    }

    if (this.isComponentIncluded('recordTypeVisibilities')) {
      this.addRecordTypeVisibilities(
        sheet,
        getMetadataPropAsArray('recordTypeVisibilities', metadataRecord),
        isProfile
      );
    }

    if (isProfile && this.isComponentIncluded('layoutAssignments')) {
      this.addPageLayoutAssignments(sheet, getMetadataPropAsArray('layoutAssignments', metadataRecord));
    }

    if (this.isComponentIncluded('objectPermissions')) {
      this.addObjectPermissions(sheet, getMetadataPropAsArray('objectPermissions', metadataRecord));
    }

    if (this.isComponentIncluded('fieldPermissions')) {
      this.addFieldPermissions(
        sheet,
        getMetadataPropAsArray('fieldPermissions', metadataRecord),
        getMetadataPropAsArray('objectPermissions', metadataRecord)
      );
    }

    if (this.isComponentIncluded('tabSettings')) {
      const propName = isProfile ? 'tabVisibilities' : 'tabSettings';
      this.addTabVisibilities(sheet, getMetadataPropAsArray(propName, metadataRecord));
    }

    if (this.isComponentIncluded('classAccesses')) {
      this.addApexClassAccesses(sheet, getMetadataPropAsArray('classAccesses', metadataRecord));
    }

    if (this.isComponentIncluded('pageAccesses')) {
      this.addPageAccesses(sheet, getMetadataPropAsArray('pageAccesses', metadataRecord));
    }

    if (this.isComponentIncluded('flowAccesses')) {
      this.addFlowAccesses(sheet, getMetadataPropAsArray('flowAccesses', metadataRecord));
    }

    if (this.isComponentIncluded('customPermissions')) {
      this.addCustomPermissions(sheet, getMetadataPropAsArray('customPermissions', metadataRecord));
    }

    if (this.isComponentIncluded('customMetadataTypeAccesses')) {
      this.addCustomMetadataTypeAccesses(sheet, getMetadataPropAsArray('customMetadataTypeAccesses', metadataRecord));
    }

    if (this.isComponentIncluded('customSettingAccesses')) {
      this.addCustomSettingAccesses(sheet, getMetadataPropAsArray('customSettingAccesses', metadataRecord));
    }

    if (this.isComponentIncluded('userPermissions')) {
      this.addUserPermissions(sheet, getMetadataPropAsArray('userPermissions', metadataRecord));
    }
  };

  private createPermissionsSheet = (workbook: Workbook, sheetName: string, sheetTitle: string): Worksheet => {
    sheetName = sheetName
      .replace('*', '')
      .replace('?', '')
      .replace(':', '')
      .replace('\\', '')
      .replace('/', '')
      .replace('[', '')
      .replace(']', '');
    sheetName = sheetName.substring(0, 31);

    let existingSheet = workbook.getWorksheet(sheetName);
    while (existingSheet) {
      sheetName = '1' + sheetName;
      sheetName = sheetName.substring(0, 31);
      existingSheet = workbook.getWorksheet(sheetName);
    }

    const sheet = workbook.addWorksheet(sheetName);

    for (let i = 1; i <= 5; i++) {
      sheet.getColumn(i).width = 50;
    }

    this.buildSheetHeader(sheet, sheetTitle);
    sheet.addRow(['']);

    return sheet;
  };

  private createPermissionSetSheets = async (permissionSets: PermissionSet[], workbook: Workbook): Promise<void> => {
    const validPermissionSetFullNames = permissionSets.map((ps) => {
      let fullName = ps.Name;
      if (ps.NamespacePrefix) {
        fullName = ps.NamespacePrefix + '__' + ps.Name;
      }
      return fullName;
    });
    this.log(`Processing the following permission sets: ${validPermissionSetFullNames.join(', ')}`);

    const permissionSetMetadataRecords = await getMetadataAsArray<ProfileOrPermissionSetMetadata>(
      this.conn,
      'PermissionSet',
      validPermissionSetFullNames
    );
    this.ux.logJson(permissionSetMetadataRecords);

    permissionSets.map((ps) => {
      this.log(`Starting building sheet for permission set: ${ps.Label}`);
      let fullName = ps.Name;
      if (ps.NamespacePrefix) {
        fullName = ps.NamespacePrefix + '__' + ps.Name;
      }

      const metadataRecord = permissionSetMetadataRecords.find((psm) => psm.fullName === fullName);
      const sheet = this.createPermissionsSheet(workbook, ps.Name, `Permission Set: ${ps.Label}`);
      this.addPermissionsToSheet(sheet, metadataRecord, false);
      this.log(`Finished building sheet for permission set: ${ps.Label}`);
    });
  };

  private createProfileSheets = async (profiles: Profile[], workbook: Workbook): Promise<void> => {
    const validProfileNames = profiles.map((ps) => ps.Name);
    this.log(`Processing the following profiles: ${validProfileNames.join(', ')}`);

    const profileMetadataRecords = await getMetadataAsArray<ProfileOrPermissionSetMetadata>(
      this.conn,
      'Profile',
      validProfileNames
    );

    profiles.map((profile) => {
      this.log(`Started building sheet for profile: ${profile.Name}`);
      const metadataRecord = profileMetadataRecords.find((mr) => mr.fullName === profile.Name);
      const sheet = this.createPermissionsSheet(workbook, profile.Name, `Profile: ${profile.Name}`);
      this.addPermissionsToSheet(sheet, metadataRecord, true);
      this.log(`Finished building sheet for profile: ${profile.Name}`);
    });
  };

  private createPermissionSetGroupSheets = async (
    permissionSetGroups: PermissionSetGroup[],
    workbook: Workbook
  ): Promise<void> => {
    const validPermissionSetGroupNames = permissionSetGroups.map((psg) => psg.DeveloperName);
    this.log(`Processing the following permission set groups: ${validPermissionSetGroupNames.join(', ')}`);

    const psgMetadataRecords = await getMetadataAsArray<PermissionSetGroupMetadata>(
      this.conn,
      'PermissionSetGroup',
      validPermissionSetGroupNames
    );

    await Promise.all(
      permissionSetGroups.map(async (psg) => {
        this.log(`Started building sheet for permission set group: ${psg.MasterLabel}`);
        const psgMetadataRecord = psgMetadataRecords.find((mr) => mr.fullName === psg.DeveloperName);
        const combinedPermissions = await new PermissionSetGroupCombine(psgMetadataRecord, this.conn).run();
        const sheet = this.createPermissionsSheet(
          workbook,
          psg.MasterLabel,
          `Permission Set Group: ${psg.MasterLabel}`
        );
        this.addPermissionsToSheet(sheet, combinedPermissions, false);
        this.log(`Finished building sheet for permission set group: ${psg.MasterLabel}`);
      })
    );
  };

  private createUserSheets = async (users: User[], workbook: Workbook): Promise<void> => {
    const validUserNames = users.map((u) => u.Username);
    const validUserIds = users.map((u) => u.Id);
    this.log(`Processing the following users: ${validUserNames.join(', ')}`);

    const permissionSetAssignments = await this.queryPermissionSetAssignments(validUserIds);
    await Promise.all(
      users.map(async (u) => {
        this.log(`Started building sheet for user: ${u.Name}`);
        const profileMetadataRecords = await getMetadataAsArray<ProfileOrPermissionSetMetadata>(this.conn, 'Profile', [
          u.Profile.Name,
        ]);

        const permissionSetNames = permissionSetAssignments
          .filter((psa) => psa.PermissionSetId && psa.AssigneeId === u.Id)
          .map((psa) => {
            let fullName = psa.PermissionSet.Name;
            if (psa.PermissionSet.NamespacePrefix) {
              fullName = psa.PermissionSet.NamespacePrefix + '__' + psa.PermissionSet.Name;
            }
            return fullName;
          });

        let permissionSetMetadataRecords: ProfileOrPermissionSetMetadata[] = [];
        if (permissionSetNames?.length) {
          permissionSetMetadataRecords = await getMetadataAsArray<ProfileOrPermissionSetMetadata>(
            this.conn,
            'PermissionSet',
            permissionSetNames
          );
        }

        const permissionSetGroupNames = permissionSetAssignments
          .filter((psa) => psa.PermissionSetGroupId && psa.AssigneeId === u.Id)
          .map((psa) => psa.PermissionSetGroup.DeveloperName);

        let psgCombinedPermissionRecords: ProfileOrPermissionSetMetadata[] = [];
        if (permissionSetGroupNames?.length) {
          const psgMetadataRecords = await getMetadataAsArray<PermissionSetGroupMetadata>(
            this.conn,
            'PermissionSetGroup',
            permissionSetGroupNames
          );

          psgCombinedPermissionRecords = await Promise.all(
            psgMetadataRecords.map(async (psgMetadataRecord) => {
              const combined = await new PermissionSetGroupCombine(psgMetadataRecord, this.conn).run();
              return combined;
            })
          );
        }

        const combinedUserPermissions = new UserPermissionsCombine(profileMetadataRecords[0], [
          ...permissionSetMetadataRecords,
          ...psgCombinedPermissionRecords,
        ]).run();

        const sheet = this.createPermissionsSheet(workbook, u.Username, `User: ${u.Name}`);
        this.addPermissionsToSheet(sheet, combinedUserPermissions, true);
        this.log(`Finished building sheet for user: ${u.Name}`);
      })
    );
  };
}
