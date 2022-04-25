import * as os from 'os';
import { flags, SfdxCommand } from '@salesforce/command';
import { Messages, SfdxError } from '@salesforce/core';
import { AnyJson } from '@salesforce/ts-types';
import PermissionsExportBuilder from '../../../services/permissions/permissionsExportBuilder';

// Initialize Messages with the current plugin directory
Messages.importMessagesDirectory(__dirname);

// Load the specific messages for this file. Messages from @salesforce/command, @salesforce/core,
// or any library that is using the messages framework can also be loaded this way.
const messages = Messages.loadMessages('sfdx-candoris-plugin', 'export');

export default class Export extends SfdxCommand {
  public static description = messages.getMessage('commandDescription');

  public static examples = messages.getMessage('examples').split(os.EOL);

  public static args = [{ name: 'file' }];

  protected static flagsConfig = {
    filename: flags.string({
      char: 'n',
      default: 'permissions',
      description: messages.getMessage('fileNameFlagDescription'),
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
    if (!permissionSetNames?.length && !profileNames?.length) {
      throw new SfdxError('Permission set names or profile names must be provided.');
    }

    const conn = this.org.getConnection();
    this.ux.startSpinner('Building permissions spreadsheet');
    const permissionsExportBuilder = new PermissionsExportBuilder(conn);
    await permissionsExportBuilder.generatePermissionsXLS(permissionSetNames, profileNames);
    this.ux.stopSpinner('Finished building permissions spreadsheet');

    return { outputString: 'Finished building permissions spreadsheet' };
  }
}
