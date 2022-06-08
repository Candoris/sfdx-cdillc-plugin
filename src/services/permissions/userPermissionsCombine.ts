import { getMetadataPropAsArray } from '../../utils';
import {
  ProfileOrPermissionSetMetadata,
  ApplicationVisibility,
  ApexClassAccess,
  CustomMetadataTypeAccess,
  CustomPermission,
  CustomSettingAccess,
  FieldPermission,
  FlowAccess,
  ObjectPermission,
  ApexPageAccess,
  RecordTypeVisibility,
  TabSetting,
  UserPermission,
} from './permissionsModels';

export default class UserPermissionsCombine {
  private profileMetadata: ProfileOrPermissionSetMetadata;
  private permissionSetMetadata: ProfileOrPermissionSetMetadata[];

  public constructor(
    profileMetadata: ProfileOrPermissionSetMetadata,
    permissionSetMetadata: ProfileOrPermissionSetMetadata[]
  ) {
    this.profileMetadata = profileMetadata;
    this.permissionSetMetadata = permissionSetMetadata;
    this.permissionSetMetadata.push(profileMetadata);
  }

  public run = (): ProfileOrPermissionSetMetadata => {
    const combinedPermissions = {} as ProfileOrPermissionSetMetadata;

    combinedPermissions.layoutAssignments = this.profileMetadata.layoutAssignments;

    if (this.permissionSetMetadata.length) {
      combinedPermissions.applicationVisibilities = this.getCombinedApplicationVisibilities();
      combinedPermissions.classAccesses = this.getCombinedApexClassAccesses();
      combinedPermissions.customMetadataTypeAccesses = this.getCombinedCustomMetadataTypeAccesses();
      combinedPermissions.customPermissions = this.getCombinedCustomPermissions();
      combinedPermissions.customSettingAccesses = this.getCombinedCustomSettingAccesses();
      combinedPermissions.fieldPermissions = this.getCombinedFieldPermissions();
      combinedPermissions.flowAccesses = this.getCombinedFlowAccesses();
      combinedPermissions.objectPermissions = this.getCombinedObjectPermissions();
      combinedPermissions.pageAccesses = this.getCombinedVisualforcePageAccesses();
      combinedPermissions.recordTypeVisibilities = this.getCombinedRecordTypeVisibilities();
      combinedPermissions.tabVisibilities = this.getCombinedTabSettings();
      combinedPermissions.userPermissions = this.getCombinedUserPermissions();
    }

    return combinedPermissions;
  };

  private getCombinedApplicationVisibilities = (): ApplicationVisibility[] => {
    const metadataMap = new Map<string, ApplicationVisibility>();

    this.permissionSetMetadata.forEach((ps) => {
      const metadataList: ApplicationVisibility[] = getMetadataPropAsArray('applicationVisibilities', ps);
      if (metadataList?.length) {
        metadataList.forEach((current) => {
          const existing = metadataMap.get(current.application);
          if (!existing) {
            metadataMap.set(current.application, current);
          } else {
            if (current.visible === 'true' && existing.visible === 'false') {
              existing.visible = 'true';
            }
            if (current.default === 'true' && existing.default === 'false') {
              existing.default = 'true';
            }
          }
        });
      }
    });

    return metadataMap.size > 0 ? [...metadataMap.values()] : null;
  };

  private getCombinedApexClassAccesses = (): ApexClassAccess[] => {
    const metadataMap = new Map<string, ApexClassAccess>();

    this.permissionSetMetadata.forEach((ps) => {
      const metadataList: ApexClassAccess[] = getMetadataPropAsArray('classAccesses', ps);
      if (metadataList?.length) {
        metadataList.forEach((current) => {
          const existing = metadataMap.get(current.apexClass);
          if (!existing) {
            metadataMap.set(current.apexClass, current);
          } else {
            if (current.enabled === 'true' && existing.enabled === 'false') {
              existing.enabled = 'true';
            }
          }
        });
      }
    });

    return metadataMap.size > 0 ? [...metadataMap.values()] : null;
  };

  private getCombinedVisualforcePageAccesses = (): ApexPageAccess[] => {
    const metadataMap = new Map<string, ApexPageAccess>();

    this.permissionSetMetadata.forEach((ps) => {
      const metadataList: ApexPageAccess[] = getMetadataPropAsArray('pageAccesses', ps);
      if (metadataList?.length) {
        metadataList.forEach((current) => {
          const existing = metadataMap.get(current.apexPage);
          if (!existing) {
            metadataMap.set(current.apexPage, current);
          } else {
            if (current.enabled === 'true' && existing.enabled === 'false') {
              existing.enabled = 'true';
            }
          }
        });
      }
    });

    return metadataMap.size > 0 ? [...metadataMap.values()] : null;
  };

  private getCombinedCustomMetadataTypeAccesses = (): CustomMetadataTypeAccess[] => {
    const metadataMap = new Map<string, CustomMetadataTypeAccess>();

    this.permissionSetMetadata.forEach((ps) => {
      const metadataList: CustomMetadataTypeAccess[] = getMetadataPropAsArray('customMetadataTypeAccesses', ps);
      if (metadataList?.length) {
        metadataList.forEach((current) => {
          const existing = metadataMap.get(current.name);
          if (!existing) {
            metadataMap.set(current.name, current);
          } else {
            if (current.enabled === 'true' && existing.enabled === 'false') {
              existing.enabled = 'true';
            }
          }
        });
      }
    });

    return metadataMap.size > 0 ? [...metadataMap.values()] : null;
  };

  private getCombinedCustomSettingAccesses = (): CustomSettingAccess[] => {
    const metadataMap = new Map<string, CustomSettingAccess>();

    this.permissionSetMetadata.forEach((ps) => {
      const metadataList: CustomSettingAccess[] = getMetadataPropAsArray('customSettingAccesses', ps);
      if (metadataList?.length) {
        metadataList.forEach((current) => {
          const existing = metadataMap.get(current.name);
          if (!existing) {
            metadataMap.set(current.name, current);
          } else {
            if (current.enabled === 'true' && existing.enabled === 'false') {
              existing.enabled = 'true';
            }
          }
        });
      }
    });

    return metadataMap.size > 0 ? [...metadataMap.values()] : null;
  };

  private getCombinedCustomPermissions = (): CustomPermission[] => {
    const metadataMap = new Map<string, CustomPermission>();

    this.permissionSetMetadata.forEach((ps) => {
      const metadataList: CustomPermission[] = getMetadataPropAsArray('customPermissions', ps);
      if (metadataList?.length) {
        metadataList.forEach((current) => {
          const existing = metadataMap.get(current.name);
          if (!existing) {
            metadataMap.set(current.name, current);
          } else {
            if (current.enabled === 'true' && existing.enabled === 'false') {
              existing.enabled = 'true';
            }
          }
        });
      }
    });

    return metadataMap.size > 0 ? [...metadataMap.values()] : null;
  };

  private getCombinedFlowAccesses = (): FlowAccess[] => {
    const metadataMap = new Map<string, FlowAccess>();

    this.permissionSetMetadata.forEach((ps) => {
      const metadataList: FlowAccess[] = getMetadataPropAsArray('flowAccesses', ps);
      if (metadataList?.length) {
        metadataList.forEach((current) => {
          const existing = metadataMap.get(current.flow);
          if (!existing) {
            metadataMap.set(current.flow, current);
          } else {
            if (current.enabled === 'true' && existing.enabled === 'false') {
              existing.enabled = 'true';
            }
          }
        });
      }
    });

    return metadataMap.size > 0 ? [...metadataMap.values()] : null;
  };

  private getCombinedRecordTypeVisibilities = (): RecordTypeVisibility[] => {
    const metadataMap = new Map<string, RecordTypeVisibility>();

    this.permissionSetMetadata.forEach((ps) => {
      const metadataList: RecordTypeVisibility[] = getMetadataPropAsArray('recordTypeVisibilities', ps);
      if (metadataList?.length) {
        metadataList.forEach((current) => {
          const existing = metadataMap.get(current.recordType);
          if (!existing) {
            metadataMap.set(current.recordType, current);
          } else {
            if (current.visible === 'true' && existing.visible === 'false') {
              existing.visible = 'true';
            }
          }
        });
      }
    });

    return metadataMap.size > 0 ? [...metadataMap.values()] : null;
  };

  private getCombinedTabSettings = (): TabSetting[] => {
    const metadataMap = new Map<string, TabSetting>();

    this.permissionSetMetadata.forEach((ps) => {
      const metadataList: TabSetting[] =
        getMetadataPropAsArray('tabSettings', ps) || getMetadataPropAsArray('tabVisibilities', ps);
      if (metadataList?.length) {
        metadataList.forEach((current) => {
          const existing = metadataMap.get(current.tab);
          if (!existing) {
            metadataMap.set(current.tab, current);
          } else {
            if (existing.visibility === 'Available' && current.visibility === 'Visible') {
              existing.visibility = 'Visible';
            } else if (existing.visibility === 'None' && current.visibility === 'Available') {
              existing.visibility = 'Available';
            }
          }
        });
      }
    });

    return metadataMap.size > 0 ? [...metadataMap.values()] : null;
  };

  private getCombinedUserPermissions = (): UserPermission[] => {
    const metadataMap = new Map<string, UserPermission>();

    this.permissionSetMetadata.forEach((ps) => {
      const metadataList: UserPermission[] = getMetadataPropAsArray('userPermissions', ps);
      if (metadataList?.length) {
        metadataList.forEach((current) => {
          const existing = metadataMap.get(current.name);
          if (!existing) {
            metadataMap.set(current.name, current);
          } else {
            if (current.enabled === 'true' && existing.enabled === 'false') {
              existing.enabled = 'true';
            }
          }
        });
      }
    });

    return metadataMap.size > 0 ? [...metadataMap.values()] : null;
  };

  private getCombinedObjectPermissions = (): ObjectPermission[] => {
    const metadataMap = new Map<string, ObjectPermission>();

    this.permissionSetMetadata.forEach((ps) => {
      const metadataList: ObjectPermission[] = getMetadataPropAsArray('objectPermissions', ps);
      if (metadataList?.length) {
        metadataList.forEach((current) => {
          const existing = metadataMap.get(current.object);
          if (!existing) {
            metadataMap.set(current.object, current);
          } else {
            if (current.allowRead === 'true' && existing.allowRead === 'false') {
              existing.allowRead = 'true';
            }
            if (current.allowCreate === 'true' && existing.allowCreate === 'false') {
              existing.allowCreate = 'true';
            }
            if (current.allowEdit === 'true' && existing.allowEdit === 'false') {
              existing.allowEdit = 'true';
            }
            if (current.allowDelete === 'true' && existing.allowDelete === 'false') {
              existing.allowDelete = 'true';
            }
            if (current.viewAllRecords === 'true' && existing.viewAllRecords === 'false') {
              existing.viewAllRecords = 'true';
            }
            if (current.modifyAllRecords === 'true' && existing.modifyAllRecords === 'false') {
              existing.modifyAllRecords = 'true';
            }
          }
        });
      }
    });

    return metadataMap.size > 0 ? [...metadataMap.values()] : null;
  };

  private getCombinedFieldPermissions = (): FieldPermission[] => {
    const metadataMap = new Map<string, FieldPermission>();

    this.permissionSetMetadata.forEach((ps) => {
      const metadataList: FieldPermission[] = getMetadataPropAsArray('fieldPermissions', ps);
      if (metadataList?.length) {
        metadataList.forEach((current) => {
          const existing = metadataMap.get(current.field);
          if (!existing) {
            metadataMap.set(current.field, current);
          } else {
            if (current.readable === 'true' && existing.readable === 'false') {
              current.readable = 'true';
            }
            if (current.editable === 'true' && existing.editable === 'false') {
              current.editable = 'true';
            }
          }
        });
      }
    });

    return metadataMap.size > 0 ? [...metadataMap.values()] : null;
  };
}
