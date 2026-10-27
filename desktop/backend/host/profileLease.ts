import { fstatSync, statSync } from 'node:fs'
import { join } from 'node:path'
/** Only a descriptor for this profile's exact lock inode may reach a writer. */
export function profileLockFd(): number | undefined {
  const value=process.env.XLSX_EDITOR_PROFILE_LOCK_FD
  if (!value) return undefined // Direct service fixtures have no Tauri parent.
  const fd=Number(value)
  if (!Number.isSafeInteger(fd)||fd<3) throw Error('permission_denied: Invalid profile lock descriptor')
  const held=fstatSync(fd), expected=statSync(join(process.env.XLSX_EDITOR_USER_DATA!,'.tauri-profile.lock'))
  if (!held.isFile()||held.dev!==expected.dev||held.ino!==expected.ino) throw Error('permission_denied: Profile lock descriptor does not match this profile')
  return fd
}
