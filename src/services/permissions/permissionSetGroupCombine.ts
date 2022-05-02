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
    const combinedPermissions = {} as ProfileOrPermissionSetMetadata;

    if (this.permissionSets.length) {
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
      combinedPermissions.tabSettings = this.getCombinedTabSettings();
      combinedPermissions.userPermissions = this.getCombinedUserPermissions();
    }

    return combinedPermissions;
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

  private getCombinedApplicationVisibilities = (): ApplicationVisibility[] => {
    const metadataMap = new Map<string, ApplicationVisibility>();

    this.permissionSets.forEach((ps) => {
      const metadataList: ApplicationVisibility[] = this.getMetadataPropAsArray('applicationVisibilities', ps);
      if (metadataList?.length) {
        metadataList.forEach((current) => {
          const existing = metadataMap.get(current.application);
          if (!existing) {
            metadataMap.set(current.application, current);
          } else {
            if (current.visible === 'true' && existing.visible === 'false') {
              existing.visible = 'true';
            }
          }
        });
      }
    });

    if (this.mutingPermissionSet) {
      const mutedList: ApplicationVisibility[] = this.getMetadataPropAsArray(
        'applicationVisibilities',
        this.mutingPermissionSet
      );
      if (mutedList?.length) {
        mutedList.forEach((muted) => {
          const existing = metadataMap.get(muted.application);
          if (existing && muted.visible === 'true') {
            existing.visible = 'false';
          }
        });
      }
    }

    return metadataMap.size > 0 ? [...metadataMap.values()] : null;
  };

  private getCombinedApexClassAccesses = (): ApexClassAccess[] => {
    const metadataMap = new Map<string, ApexClassAccess>();

    this.permissionSets.forEach((ps) => {
      const metadataList: ApexClassAccess[] = this.getMetadataPropAsArray('classAccesses', ps);
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

    if (this.mutingPermissionSet) {
      const mutedList: ApexClassAccess[] = this.getMetadataPropAsArray('classAccesses', this.mutingPermissionSet);
      if (mutedList?.length) {
        mutedList.forEach((muted) => {
          const existing = metadataMap.get(muted.apexClass);
          if (existing && muted.enabled === 'true') {
            existing.enabled = 'false';
          }
        });
      }
    }

    return metadataMap.size > 0 ? [...metadataMap.values()] : null;
  };

  private getCombinedVisualforcePageAccesses = (): ApexPageAccess[] => {
    const metadataMap = new Map<string, ApexPageAccess>();

    this.permissionSets.forEach((ps) => {
      const metadataList: ApexPageAccess[] = this.getMetadataPropAsArray('pageAccesses', ps);
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

    if (this.mutingPermissionSet) {
      const mutedList: ApexPageAccess[] = this.getMetadataPropAsArray('pageAccesses', this.mutingPermissionSet);
      if (mutedList?.length) {
        mutedList.forEach((muted) => {
          const existing = metadataMap.get(muted.apexPage);
          if (existing && muted.enabled === 'true') {
            existing.enabled = 'false';
          }
        });
      }
    }

    return metadataMap.size > 0 ? [...metadataMap.values()] : null;
  };

  private getCombinedCustomMetadataTypeAccesses = (): CustomMetadataTypeAccess[] => {
    const metadataMap = new Map<string, CustomMetadataTypeAccess>();

    this.permissionSets.forEach((ps) => {
      const metadataList: CustomMetadataTypeAccess[] = this.getMetadataPropAsArray('customMetadataTypeAccesses', ps);
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

    if (this.mutingPermissionSet) {
      const mutedList: CustomMetadataTypeAccess[] = this.getMetadataPropAsArray(
        'customMetadataTypeAccesses',
        this.mutingPermissionSet
      );
      if (mutedList?.length) {
        mutedList.forEach((muted) => {
          const existing = metadataMap.get(muted.name);
          if (existing && muted.enabled === 'true') {
            existing.enabled = 'false';
          }
        });
      }
    }

    return metadataMap.size > 0 ? [...metadataMap.values()] : null;
  };

  private getCombinedCustomSettingAccesses = (): CustomSettingAccess[] => {
    const metadataMap = new Map<string, CustomSettingAccess>();

    this.permissionSets.forEach((ps) => {
      const metadataList: CustomSettingAccess[] = this.getMetadataPropAsArray('customSettingAccesses', ps);
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

    if (this.mutingPermissionSet) {
      const mutedList: CustomSettingAccess[] = this.getMetadataPropAsArray(
        'customSettingAccesses',
        this.mutingPermissionSet
      );
      if (mutedList?.length) {
        mutedList.forEach((muted) => {
          const existing = metadataMap.get(muted.name);
          if (existing && muted.enabled === 'true') {
            existing.enabled = 'false';
          }
        });
      }
    }

    return metadataMap.size > 0 ? [...metadataMap.values()] : null;
  };

  private getCombinedCustomPermissions = (): CustomPermission[] => {
    const metadataMap = new Map<string, CustomPermission>();

    this.permissionSets.forEach((ps) => {
      const metadataList: CustomPermission[] = this.getMetadataPropAsArray('customPermissions', ps);
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

    if (this.mutingPermissionSet) {
      const mutedList: CustomPermission[] = this.getMetadataPropAsArray('customPermissions', this.mutingPermissionSet);
      if (mutedList?.length) {
        mutedList.forEach((muted) => {
          const existing = metadataMap.get(muted.name);
          if (existing && muted.enabled === 'true') {
            existing.enabled = 'false';
          }
        });
      }
    }

    return metadataMap.size > 0 ? [...metadataMap.values()] : null;
  };

  private getCombinedFlowAccesses = (): FlowAccess[] => {
    const metadataMap = new Map<string, FlowAccess>();

    this.permissionSets.forEach((ps) => {
      const metadataList: FlowAccess[] = this.getMetadataPropAsArray('flowAccesses', ps);
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

    if (this.mutingPermissionSet) {
      const mutedList: FlowAccess[] = this.getMetadataPropAsArray('flowAccesses', this.mutingPermissionSet);
      if (mutedList?.length) {
        mutedList.forEach((muted) => {
          const existing = metadataMap.get(muted.flow);
          if (existing && muted.enabled === 'true') {
            existing.enabled = 'false';
          }
        });
      }
    }

    return metadataMap.size > 0 ? [...metadataMap.values()] : null;
  };

  private getCombinedRecordTypeVisibilities = (): RecordTypeVisibility[] => {
    const metadataMap = new Map<string, RecordTypeVisibility>();

    this.permissionSets.forEach((ps) => {
      const metadataList: RecordTypeVisibility[] = this.getMetadataPropAsArray('recordTypeVisibilities', ps);
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

    if (this.mutingPermissionSet) {
      const mutedList: RecordTypeVisibility[] = this.getMetadataPropAsArray(
        'recordTypeVisibilities',
        this.mutingPermissionSet
      );
      if (mutedList?.length) {
        mutedList.forEach((muted) => {
          const existing = metadataMap.get(muted.recordType);
          if (existing && muted.visible === 'true') {
            existing.visible = 'false';
          }
        });
      }
    }

    return metadataMap.size > 0 ? [...metadataMap.values()] : null;
  };

  private getCombinedTabSettings = (): TabSetting[] => {
    const metadataMap = new Map<string, TabSetting>();

    this.permissionSets.forEach((ps) => {
      const metadataList: TabSetting[] = this.getMetadataPropAsArray('recordTypeVisibilities', ps);
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

    if (this.mutingPermissionSet) {
      const mutedList: TabSetting[] = this.getMetadataPropAsArray('recordTypeVisibilities', this.mutingPermissionSet);
      if (mutedList?.length) {
        mutedList.forEach((muted) => {
          const existing = metadataMap.get(muted.tab);
          if (existing && muted.visibility === 'Available') {
            existing.visibility = 'None';
          }
        });
      }
    }

    return metadataMap.size > 0 ? [...metadataMap.values()] : null;
  };

  private getCombinedUserPermissions = (): UserPermission[] => {
    const metadataMap = new Map<string, UserPermission>();

    this.permissionSets.forEach((ps) => {
      const metadataList: UserPermission[] = this.getMetadataPropAsArray('userPermissions', ps);
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

    if (this.mutingPermissionSet) {
      const mutedList: UserPermission[] = this.getMetadataPropAsArray('userPermissions', this.mutingPermissionSet);
      if (mutedList?.length) {
        mutedList.forEach((muted) => {
          const existing = metadataMap.get(muted.name);
          if (existing && muted.enabled === 'true') {
            existing.enabled = 'false';
          }
        });
      }
    }

    return metadataMap.size > 0 ? [...metadataMap.values()] : null;
  };

  private getCombinedObjectPermissions = (): ObjectPermission[] => {
    const metadataMap = new Map<string, ObjectPermission>();

    this.permissionSets.forEach((ps) => {
      const metadataList: ObjectPermission[] = this.getMetadataPropAsArray('objectPermissions', ps);
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

    if (this.mutingPermissionSet) {
      const mutedList: ObjectPermission[] = this.getMetadataPropAsArray('objectPermissions', this.mutingPermissionSet);
      if (mutedList?.length) {
        mutedList.forEach((muted) => {
          const existing = metadataMap.get(muted.object);
          if (existing) {
            if (existing.allowRead === 'true' && muted.allowRead === 'true') {
              existing.allowRead = 'false';
            }
            if (existing.allowCreate === 'true' && muted.allowCreate === 'true') {
              existing.allowCreate = 'false';
            }
            if (existing.allowEdit === 'true' && muted.allowEdit === 'true') {
              existing.allowEdit = 'false';
            }
            if (existing.allowDelete === 'true' && muted.allowDelete === 'true') {
              existing.allowDelete = 'false';
            }
            if (existing.viewAllRecords === 'true' && muted.viewAllRecords === 'true') {
              existing.viewAllRecords = 'false';
            }
            if (existing.modifyAllRecords === 'true' && muted.modifyAllRecords === 'true') {
              existing.modifyAllRecords = 'false';
            }
          }
        });
      }
    }

    return metadataMap.size > 0 ? [...metadataMap.values()] : null;
  };

  private getCombinedFieldPermissions = (): FieldPermission[] => {
    const metadataMap = new Map<string, FieldPermission>();

    this.permissionSets.forEach((ps) => {
      const metadataList: FieldPermission[] = this.getMetadataPropAsArray('fieldPermissions', ps);
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

    if (this.mutingPermissionSet) {
      const mutedList: FieldPermission[] = this.getMetadataPropAsArray('fieldPermissions', this.mutingPermissionSet);
      if (mutedList?.length) {
        mutedList.forEach((muted) => {
          const existing = metadataMap.get(muted.field);
          if (existing) {
            if (existing.readable === 'true' && muted.readable === 'true') {
              existing.readable = 'false';
            }
            if (existing.editable === 'true' && muted.editable === 'true') {
              existing.editable = 'false';
            }
          }
        });
      }
    }

    return metadataMap.size > 0 ? [...metadataMap.values()] : null;
  };
}
