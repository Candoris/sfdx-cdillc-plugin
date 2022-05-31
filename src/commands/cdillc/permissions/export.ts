import * as os from 'os';
import { flags, SfdxCommand } from '@salesforce/command';
import { Messages, SfdxError } from '@salesforce/core';
import { AnyJson } from '@salesforce/ts-types';
import PermissionsExportBuilder from '../../../services/permissions/permissionsExportBuilder';

// Initialize Messages with the current plugin directory
Messages.importMessagesDirectory(__dirname);

// Load the specific messages for this file. Messages from @salesforce/command, @salesforce/core,
// or any library that is using the messages framework can also be loaded this way.
const messages = Messages.loadMessages('sfdx-cdillc-plugin', 'permissionsExport');

export default class Export extends SfdxCommand {
  public static description = messages.getMessage('commandDescription');

  public static examples = messages.getMessage('examples').split(os.EOL);

  public static args = [{ name: 'file' }];

  protected static flagsConfig = {
    filepath: flags.string({
      char: 'f',
      default: 'permissions.xlsx',
      description: messages.getMessage('filePathFlagDescription'),
    }),
    profilenames: flags.array({
      char: 'p',
      delimiter: ',',
      description: messages.getMessage('profileNamesFlagDescription'),
    }),
    permissionsetnames: flags.array({
      char: 's',
      delimiter: ',',
      description: messages.getMessage('permissionSetNamesFlagDescription'),
    }),
    permissionsetgroupnames: flags.array({
      char: 'g',
      delimiter: ',',
      description: messages.getMessage('permissionSetGroupNamesFlagDescription'),
    }),
    includedcomponents: flags.array({
      char: 'i',
      delimiter: ',',
      default: ['all'],
      description: messages.getMessage('includedComponentsFlagDescription'),
    }),
    verbose: flags.builtin(),
  };

  // Comment this out if your command does not require an org username
  protected static requiresUsername = true;

  // Comment this out if your command does not support a hub org username
  // protected static supportsDevhubUsername = true;

  // Set this to true if your command requires a project workspace; 'requiresProject' is false by default
  protected static requiresProject = false;

  public async run(): Promise<AnyJson> {
    const permissionSetNames = this.flags.permissionsetnames as string[];
    const profileNames = this.flags.profilenames as string[];
    const permissionSetGroupNames = this.flags.permissionsetgroupnames as string[];
    if (!permissionSetNames?.length && !profileNames?.length && !permissionSetGroupNames?.length) {
      throw new SfdxError('Permission set names, profile names, or permission set group names must be provided.');
    }
    const includedComponents = this.flags.includedcomponents as string[];
    const verbose = this.flags.verbose as boolean;

    const conn = this.org.getConnection();
    this.ux.startSpinner('Building permissions spreadsheet');
    const permissionsExportBuilder = new PermissionsExportBuilder(conn, this.ux, includedComponents, verbose);
    await permissionsExportBuilder.generatePermissionsXLS(
      permissionSetNames,
      profileNames,
      permissionSetGroupNames,
      this.flags.filepath
    );
    this.ux.stopSpinner();

    return;
  }
}
