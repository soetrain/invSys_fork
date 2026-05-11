## Synology DS920+ invSys folders
| Folder         | invsys-svc (Codex) | justin-invsys (you) | Personal account |
| -------------- | ------------------ | ------------------- | ---------------- |
| invSysWH1      | Read only          | Read/Write          | No access        |
| invSys-deploy  | Read/Write         | Read/Write          | No access        |
| invSys-backups | Read only          | Read only           | No access        |

## invsys-agents — a file-access-only service group for invSys.
invsys-agents Application Permissions:
| Application                       | Setting | Reason                                                                                                                                 |
| --------------------------------- | ------- | -------------------------------------------------------------------------------------------------------------------------------------- |
| SMB                               | ✅ Allow | LAN stations and processor PC access NAS shares over SMB — required                                                                    |
| SFTP                              | ✅ Allow | Codex/GitHub Actions pushes built artifacts over SFTP — required                                                                       |
| Synology Drive                    | ✅ Allow | Your remote file access and optional relay sync path                                                                                   |
| File Station                      | ✅ Allow | You need browser-based file access when working remotely                                                                               |
| rsync                             | ✅ Allow | Needed if you use Shared Folder Sync for backups or relay                                                                              |
| FTP                               | ⛔ Deny  | Not needed — SFTP covers secure file transfer; plain FTP is unencrypted                                                                |
| AFP                               | ⛔ Deny  | Apple Filing Protocol — legacy Mac protocol, not needed                                                                                |
| Note Station                      | ⛔ Deny  | Personal productivity app, not relevant to invSys                                                                                      |
| Synology Photos                   | ⛔ Deny  | Personal media app, not relevant                                                                                                       |
| Universal Search                  | ⛔ Deny  | DSM search UI feature, not needed for service accounts                                                                                 |
| Audio Station                     | ⛔ Deny  | Personal media app, not relevant                                                                                                       |
| Active Backup for Business        | ⛔ Deny  | Backup agent management — admin-only function                                                                                          |
| Active Backup for Business Portal | ⛔ Deny  | Same — admin only                                                                                                                      |
| Active Backup for Business Agent  | ⛔ Deny  | Same — admin only                                                                                                                      |
| DSM                               | ⛔ Deny  | Neither invsys-svc nor your justin-invsys working account needs DSM admin UI access — manage DSM from your personal admin account only |


## Two-Account Split After Group Creation
**`invsys-svc` (Codex service account):**
- Deny **File Station** and **Synology Drive** too — Codex only needs SFTP, not a browser UI
- This minimizes the attack surface if the credentials are ever compromised

**`justin-invsys` (your invSys working account):**
- Keep File Station and Synology Drive allowed — you need those for remote review and manual promotion of Codex builds to `invSysWH1`

## settings for invSysWH1
Local Group:
| Name           | Setting                        | Reason                                                                               |
| -------------- | ------------------------------ | ------------------------------------------------------------------------------------ |
| administrators | ✅ Read/Write (already checked) | Keep — admin needs full access for maintenance                                       |
| http           | ⛔ No Access                    | Keep as-is — web server process has no business here                                 |
| invsys-agents  | ✅ Read Only                    | The group gets read access; individual user overrides below grant write where needed |
| users          | ⛔ No Access                    | Keep as-is — default user group should not see warehouse files                       |

Local Users:
| Name          | Setting                      | Reason                                                                                       |
| ------------- | ---------------------------- | -------------------------------------------------------------------------------------------- |
| admin         | Read/Write (inherited, keep) | DSM built-in admin                                                                           |
| Git NAS repo  | ⛔ No Access                  | Not relevant to invSysWH1                                                                    |
| guest         | ⛔ No Access                  | Keep as-is                                                                                   |
| invsys-justin | ✅ Read/Write                 | Your working account — you need to promote Codex builds and manage files                     |
| invsys-svc    | ✅ Read Only                  | Codex service account — can read canonical files, cannot write or delete                     |
| justinwj      | ⛔ No Access                  | Your personal account should not touch warehouse runtime — use invsys-justin for invSys work |

## invSys-deploy is the staging area where Codex pushes built .xlam artifacts
Local Groups:
| Name           | Setting      | Reason                                                                                 |
| -------------- | ------------ | -------------------------------------------------------------------------------------- |
| administrators | ✅ Read/Write | Keep — admin needs full access                                                         |
| http           | ⛔ No Access  | Not needed                                                                             |
| invsys-agents  | ✅ Read/Write | Both Codex and your working account are in this group — deploy folder is safe to write |
| users          | ⛔ No Access  | Keep as-is                                                                             |

Local Users:
| Name          | Setting           | Reason                                                                  |
| ------------- | ----------------- | ----------------------------------------------------------------------- |
| admin         | Read/Write (keep) | DSM built-in admin                                                      |
| Git NAS repo  | ⛔ No Access       | Not relevant                                                            |
| guest         | ⛔ No Access       | Keep as-is                                                              |
| invsys-justin | ✅ Read/Write      | You review and promote Codex builds from here                           |
| invsys-svc    | ✅ Read/Write      | Codex must write here — this is its only write-access folder on the NAS |
| justinwj      | ⛔ No Access       | Personal account stays out of invSys folders                            |

## invSys-backups is the most locked down
Local Groups:
| Name           | Setting      | Reason                                                                                    |
| -------------- | ------------ | ----------------------------------------------------------------------------------------- |
| administrators | ✅ Read/Write | Keep — admin needs access for backup management                                           |
| http           | ⛔ No Access  | Not needed                                                                                |
| invsys-agents  | ✅ Read Only  | Group can read backups but nobody in the group should be able to delete or overwrite them |
| users          | ⛔ No Access  | Keep as-is                                                                                |

Local Users:
| Name          | Setting           | Reason                                                                      |
| ------------- | ----------------- | --------------------------------------------------------------------------- |
| admin         | Read/Write (keep) | DSM built-in admin                                                          |
| Git NAS repo  | ⛔ No Access       | Not relevant                                                                |
| guest         | ⛔ No Access       | Keep as-is                                                                  |
| invsys-justin | ✅ Read Only       | You can browse and restore from backups but not accidentally overwrite them |
| invsys-svc    | ⛔ No Access       | Codex has no business reading or writing backups directly                   |
| justinwj      | ⛔ No Access       | Personal account stays out                                                  |


