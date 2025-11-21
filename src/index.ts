import { writeFileSync } from "node:fs"
import type { User } from "@microsoft/microsoft-graph-types"
import { logger } from "@vestfoldfylke/loglady"
import { DOCUMASTER_LIST, GROUP_MAPPINGS } from "./config.js"
import { addMemberToGroup, getGroupMembers, getListItems, getUserByUpn, removeMemberFromGroup } from "./lib/graph.js"
import type { DocumasterGraphListItem } from "./types/graph.types"

const logErrorAndExit = async (message: string, error: unknown) => {
	logger.errorException(error, message)
	await logger.flush()
	process.exit(1)
}

logger.info("Starting Documaster user sync process")

let documasterAccessListItems: DocumasterGraphListItem[] = []
try {
	documasterAccessListItems = (await getListItems(DOCUMASTER_LIST.SITE_ID, DOCUMASTER_LIST.LIST_ID)) as DocumasterGraphListItem[]
	if (!documasterAccessListItems || documasterAccessListItems.length === 0) {
		throw new Error("No items found in Documaster access list - make sure you have the correct SITE_ID and LIST_ID configured")
	}
} catch (error) {
	await logErrorAndExit("Failed to fetch Documaster access list items from Sharepoint", error)
}

logger.info("Fetched {itemCount} items from Documaster access list", documasterAccessListItems.length)
const totalGroups = Object.keys(GROUP_MAPPINGS).length
logger.info("Processing {totalGroups} groups based on GROUP_MAPPINGS configuration", totalGroups)

if (documasterAccessListItems.length !== totalGroups) {
	logger.warn("Number of items in Documaster access list ({itemCount}) does not match number of configured groups in GROUP_MAPPINGS ({totalGroups})", documasterAccessListItems.length, totalGroups)
}

writeFileSync("./ignore/debug-documaster-list-items.json", JSON.stringify(documasterAccessListItems, null, 2))

for (const item of documasterAccessListItems) {
	// Hent ut tilsvarende AzureADGroupId fra GROUP_MAPPINGS basert på SharePointGroupName (item.fields.Title)
	const sharePointGroupName = item.fields.Title
	const entraGroupId = GROUP_MAPPINGS[sharePointGroupName]

	if (!entraGroupId) {
		logger.warn(`No AzureADGroupId mapping found for SharePoint group name: {sharePointGroupName}. Skipping item with ID: {itemId}`, sharePointGroupName, item.id)
		continue
	}

	logger.logConfig({
		prefix: `ProcessingGroup - ${sharePointGroupName} - ${entraGroupId}`
	})

	const groupResult = {
		added: 0,
		failedToAdd: 0,
		removed: 0,
		failedToRemove: 0
	}

	logger.info("Starting processing of group, fetching members from Entra group {entraGroupId}", entraGroupId)
	// Hent medlemmer i Entra-gruppen
	let entraGroupMembers: User[] = []
	try {
		entraGroupMembers = await getGroupMembers(entraGroupId)
	} catch (error) {
		logErrorAndExit(`Failed to fetch members for Entra group with ID: ${entraGroupId}`, error)
	}
	logger.info("Fetched {memberCount} members from Entra group {entraGroupId}", entraGroupMembers.length, entraGroupId)
	logger.info("Finding users to add and remove")
	const usersToAdd = item.fields.Hartilgang.filter((spUser) => {
		return !entraGroupMembers.some((entraUser) => entraUser.mail?.toLowerCase() === spUser.Email.toLowerCase())
	})
	const usersToRemove = entraGroupMembers.filter((entraUser) => {
		return !item.fields.Hartilgang.some((spUser) => spUser.Email.toLowerCase() === entraUser.mail?.toLowerCase())
	})
	logger.info("Found {count} users in Sharepoint list to add to Entra group {entraGroupId}", usersToAdd.length, entraGroupId)
	logger.info("Found {count} users in Entra group {entraGroupId} to remove (not in Sharepoint list)", usersToRemove.length, entraGroupId)

	// Legg til brukere som mangler i Entra-gruppen
	for (const user of usersToAdd) {
		logger.info("Adding user {email} to Entra group {entraGroupId}", user.Email, entraGroupId)
		try {
			const userByUpn = await getUserByUpn(user.Email)
			if (!userByUpn) {
				logger.warn("User with email {email} not found in Entra. Skipping addition.", user.Email)
				// TODO - varsle arkivet om at bruker må fjernes fra Sharepoint-listen (eller sjekkes opp)
				groupResult.failedToAdd++
				continue
			}
			if (!userByUpn.id) {
				logger.warn("User with email {email} has no ID in Entra. WHAAAT? Skipping addition.", user.Email)
				groupResult.failedToAdd++
				continue
			}
			await addMemberToGroup(entraGroupId, userByUpn.id)
			groupResult.added++
			logger.info("Successfully added user {email} to Entra group {entraGroupId}", user.Email, entraGroupId)
		} catch (error) {
			groupResult.failedToAdd++
			logger.errorException(error, "Failed to add user {email} to Entra group {entraGroupId}", user.Email, entraGroupId)
		}
	}

	// Fjern brukere som ikke lenger skal være i Entra-gruppen
	for (const user of usersToRemove) {
		logger.info("Removing user {email} from Entra group {entraGroupId}", user.mail, entraGroupId)
		try {
			if (!user.id) {
				logger.warn("User with email {email} has no ID in Entra. WHAAAT? Skipping removal.", user.mail)
				groupResult.failedToRemove++
				continue
			}
			await removeMemberFromGroup(entraGroupId, user.id)
			groupResult.removed++
			logger.info("Successfully removed user {email} from Entra group {entraGroupId}", user.mail, entraGroupId)
		} catch (error) {
			groupResult.failedToRemove++
			logger.errorException(error, "Failed to remove user {email} from Entra group {entraGroupId}", user.mail, entraGroupId)
		}
	}

	// Logg resultatet for nåværende gruppe
	logger.info("Finished processing group {sharePointGroupName}. Summary: {@summary}", sharePointGroupName, groupResult)
}

logger.info("All groups processed, flushing logs and exiting.")
await logger.flush()
process.exit(0)
