import { UX } from '@salesforce/command';
import { Connection } from '@salesforce/core';
import { Workbook, Worksheet } from 'exceljs';
import { DescribeGlobalResult, DescribeGlobalSObjectResult, DescribeSObjectResult } from 'jsforce';
import {
  AppMenuItem,
  RecordType,
  PermissionSet,
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
} from './permissionsModels';

export default class PermissionSetGroupCombine {
  private permissionSets: ProfileOrPermissionSetMetadata[];
  private mutingPermissionSet: ProfileOrPermissionSetMetadata;

  public constructor(
    permissionSets: ProfileOrPermissionSetMetadata[],
    mutingPermissionSet: ProfileOrPermissionSetMetadata
  ) {
    this.permissionSets = permissionSets;
    this.mutingPermissionSet = mutingPermissionSet;
  }

  public run = (): ProfileOrPermissionSetMetadata => {
    const combinedPermissions = this.permissionSets[0];
    this.permissionSets.splice(0, 1);

    if (this.permissionSets.length) {
      this.permissionSets.forEach((ps) => { });
    }

    return combinedPermissions;
  };
}
