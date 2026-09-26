import { App, Modal, Setting, TFolder } from 'obsidian'
import type { SharedFolderConfig } from '@obsidian-teams/shared'
import type ObsidianTeamsPlugin from '../main'
import { decodeAccessToken, redeemInvite, storeAccessToken, storeRefreshToken } from '../utils/auth'
import { readSharedConfigAsync, writeSharedConfig } from '../utils/dotfile'

export interface JoinSharedFolderResult {
  folderId: string
  folderName: string
}

function normalizeFolderPath(value: string): string {
  return value
    .split('/')
    .map((segment) => segment.trim())
    .filter(Boolean)
    .join('/')
}

function addFolderSuffix(path: string, suffix: number): string {
  if (suffix <= 0) return path

  const lastSlash = path.lastIndexOf('/')
  const parent = lastSlash >= 0 ? path.slice(0, lastSlash) : ''
  const leaf = lastSlash >= 0 ? path.slice(lastSlash + 1) : path
  const suffixedLeaf = `${leaf} (${suffix})`
  return parent ? `${parent}/${suffixedLeaf}` : suffixedLeaf
}

async function resolveJoinTargetPath(
  app: App,
  preferredPath: string,
  expectedFolderId: string
): Promise<string> {
  const normalizedPreferredPath = normalizeFolderPath(preferredPath)
  if (!normalizedPreferredPath) {
    throw new Error('Invite response missing folder name')
  }

  for (let suffix = 0; suffix < 1_000; suffix += 1) {
    const candidatePath = addFolderSuffix(normalizedPreferredPath, suffix)
    const existing = app.vault.getAbstractFileByPath(candidatePath)
    if (!existing) {
      return candidatePath
    }
    if (!(existing instanceof TFolder)) {
      continue
    }

    const existingConfig = await readSharedConfigAsync(app.vault, candidatePath)
    if (existingConfig?.folderId === expectedFolderId) {
      return candidatePath
    }
  }

  throw new Error(`Unable to pick a unique folder path for '${normalizedPreferredPath}'`)
}

async function ensureFolderHierarchy(app: App, targetPath: string): Promise<void> {
  const segments = normalizeFolderPath(targetPath).split('/').filter(Boolean)
  let currentPath = ''

  for (const segment of segments) {
    currentPath = currentPath ? `${currentPath}/${segment}` : segment
    const existing = app.vault.getAbstractFileByPath(currentPath)
    if (!existing) {
      await app.vault.createFolder(currentPath)
      continue
    }
    if (!(existing instanceof TFolder)) {
      throw new Error(`Path '${currentPath}' already exists and is not a folder`)
    }
  }
}

async function assertNoSharedFolderConflict(
  app: App,
  targetPath: string,
  expectedFolderId: string
): Promise<void> {
  const existingConfig = await readSharedConfigAsync(app.vault, targetPath)
  if (existingConfig && existingConfig.folderId !== expectedFolderId) {
    throw new Error('Target path already contains a different shared folder')
  }
}

export interface ParsedInviteInput {
  inviteToken: string
  serverUrl?: string
}

export function parseInviteInput(input: string): ParsedInviteInput {
  const trimmed = input.trim()
  if (!trimmed) {
    return { inviteToken: '' }
  }

  // Handle obsidian://teams-join?token=...&server=...
  if (trimmed.startsWith('obsidian://')) {
    try {
      const url = new URL(trimmed)
      const token = url.searchParams.get('token') || url.searchParams.get('inviteToken') || ''
      const server = url.searchParams.get('server') || url.searchParams.get('serverUrl') || undefined
      if (token) {
        return {
          inviteToken: decodeURIComponent(token).trim(),
          serverUrl: server ? decodeURIComponent(server).trim().replace(/\/+$/, '') : undefined,
        }
      }
    } catch {
      // Fall through to regex/raw token
    }
  }

  // Handle http(s)://.../api/invite/redeem?token=...
  if (/^https?:\/\//i.test(trimmed)) {
    try {
      const url = new URL(trimmed)
      const token = url.searchParams.get('token') || url.searchParams.get('inviteToken') || ''
      const serverParam = url.searchParams.get('server') || url.searchParams.get('serverUrl')
      const serverUrl = serverParam
        ? decodeURIComponent(serverParam).trim().replace(/\/+$/, '')
        : `${url.protocol}//${url.host}`
      if (token) {
        return {
          inviteToken: decodeURIComponent(token).trim(),
          serverUrl,
        }
      }
    } catch {
      // Fall through to raw token
    }
  }

  return { inviteToken: trimmed }
}

export async function joinSharedFolderByInvite(
  app: App,
  plugin: ObsidianTeamsPlugin,
  inviteInput: string,
  options: { serverUrl?: string } = {}
): Promise<JoinSharedFolderResult> {
  const parsed = parseInviteInput(inviteInput)
  const normalizedInviteToken = parsed.inviteToken
  if (!normalizedInviteToken) {
    throw new Error('Please enter an invite token')
  }

  const effectiveServerUrl = options.serverUrl || parsed.serverUrl || plugin.settings.serverUrl
  const { clientId, displayName } = plugin.settings
  const hostedSessionToken =
    plugin.settings.deploymentMode === 'hosted-service'
      ? plugin.settings.hostedSessionToken || undefined
      : undefined

  const result = await redeemInvite(
    effectiveServerUrl,
    normalizedInviteToken,
    clientId,
    displayName || 'Anonymous',
    hostedSessionToken
  )

  await storeAccessToken(plugin, result.folderId, result.accessToken)
  await storeRefreshToken(plugin, result.folderId, result.refreshToken)

  const targetPath = await resolveJoinTargetPath(app, result.folderName, result.folderId)

  await ensureFolderHierarchy(app, targetPath)
  await assertNoSharedFolderConflict(app, targetPath, result.folderId)

  const targetFolder = app.vault.getAbstractFileByPath(targetPath)
  if (!targetFolder || !(targetFolder instanceof TFolder)) {
    throw new Error(`Unable to create target folder at '${targetPath}'`)
  }

  const access = decodeAccessToken(result.accessToken)
  const config: SharedFolderConfig = {
    folderId: result.folderId,
    serverUrl: result.serverUrl || effectiveServerUrl,
    displayName: result.folderName,
    members: [
      {
        clientId,
        name: displayName || 'Anonymous',
        role: access?.role || 'editor',
      },
    ],
    createdAt: new Date().toISOString(),
  }
  await writeSharedConfig(app.vault, targetPath, config)

  await plugin.refreshSharedFolders()

  return {
    folderId: result.folderId,
    folderName: result.folderName,
  }
}

export class JoinFolderModal extends Modal {
  private plugin: ObsidianTeamsPlugin
  private inviteToken = ''
  private serverUrl?: string
  private actionInFlight = false
  private errorMessage = ''

  constructor(app: App, plugin: ObsidianTeamsPlugin, options: { inviteToken?: string; serverUrl?: string } = {}) {
    super(app)
    this.plugin = plugin
    this.inviteToken = options.inviteToken?.trim() || ''
    this.serverUrl = options.serverUrl?.trim() || undefined
  }

  onOpen() {
    this.render()
  }

  private render() {
    const { contentEl } = this
    contentEl.empty()
    contentEl.addClass('obsidian-teams-join-modal')

    contentEl.createEl('h2', { text: 'Join shared folder' })
    contentEl.createEl('p', {
      text: 'Paste the invite link or token you received from the folder owner.',
      cls: 'setting-item-description',
    })

    if (this.errorMessage) {
      contentEl.createEl('p', {
        cls: 'setting-item-description obsidian-teams-join-error',
        text: this.errorMessage,
      })
    }

    new Setting(contentEl)
      .setName('Invite link or token')
      .addText((text) => {
        text
          .setPlaceholder('Paste invite link or token here...')
          .setValue(this.inviteToken)
          .setDisabled(this.actionInFlight)
          .onChange((value) => {
            this.inviteToken = value.trim()
          })
      })

    new Setting(contentEl).addButton((btn) => {
      btn
        .setButtonText(this.actionInFlight ? 'Joining...' : 'Join folder')
        .setCta()
        .setDisabled(this.actionInFlight || !this.inviteToken.trim())
        .onClick(async () => {
          await this.joinFolder()
        })
    })
  }

  private async joinFolder() {
    if (this.actionInFlight) return

    const input = this.inviteToken.trim()
    if (!input) {
      this.errorMessage = 'Please enter an invite token or link'
      this.render()
      return
    }

    this.actionInFlight = true
    this.errorMessage = ''
    this.render()

    try {
      const joined = await this.plugin.attemptInviteJoin(input, { serverUrl: this.serverUrl })
      if (joined) {
        this.close()
      }
    } catch (error) {
      this.errorMessage = error instanceof Error ? error.message : 'Unknown error'
      this.render()
    } finally {
      this.actionInFlight = false
      if (this.contentEl.isConnected) {
        this.render()
      }
    }
  }

  onClose() {
    const { contentEl } = this
    contentEl.empty()
  }
}
