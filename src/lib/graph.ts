import { DefaultAzureCredential } from "@azure/identity"
import type { ListItem, User } from "@microsoft/microsoft-graph-types"

if (!process.env.AZURE_TENANT_ID || !process.env.AZURE_CLIENT_ID || !process.env.AZURE_CLIENT_SECRET) {
	throw new Error("AZURE_TENANT_ID, AZURE_CLIENT_ID, and AZURE_CLIENT_SECRET must be set in environment variables")
}

const defaultAzureCredential = new DefaultAzureCredential({})

const graphScope = "https://graph.microsoft.com/.default"

async function graphGet<T>(endpoint: string): Promise<T> {
	const accessToken = await defaultAzureCredential.getToken(graphScope)
	const response = await fetch(`https://graph.microsoft.com/v1.0/${endpoint}`, { headers: { Authorization: `Bearer ${accessToken?.token}` } })

	if (!response.ok) {
		const errorData = await response.json()
		throw new Error(`Failed to fetch endpoint '${endpoint}' : ${response.status} ${response.statusText} - ${JSON.stringify(errorData, null, 2)}`)
	}

	return (await response.json()) as T
}

export const getListItems = async (siteId: string, listId: string): Promise<ListItem[]> => {
	return (await graphGet<{ value: ListItem[] }>(`sites/${siteId}/lists/${listId}/items?expand=fields`)).value
}

export const getGroupMembers = async (groupId: string): Promise<User[]> => {
	return (await graphGet<{ value: User[] }>(`groups/${groupId}/members`)).value
}

export const getUserByUpn = async (email: string): Promise<User | null> => {
	const result = await graphGet<{ value: User[] }>(`users?$filter=mail eq '${email}' or userPrincipalName eq '${email}'`)
	return result.value.length > 0 && result.value[0] !== undefined ? result.value[0] : null
}

export const addMemberToGroup = async (groupId: string, userId: string): Promise<void> => {
	const accessToken = await defaultAzureCredential.getToken(graphScope)
	const response = await fetch(`https://graph.microsoft.com/v1.0/groups/${groupId}/members/$ref`, {
		method: "POST",
		headers: {
			Authorization: `Bearer ${accessToken?.token}`,
			"Content-Type": "application/json"
		},
		body: JSON.stringify({
			"@odata.id": `https://graph.microsoft.com/v1.0/users/${userId}`
		})
	})

	if (!response.ok) {
		const errorData = await response.json()
		throw new Error(`Failed to add user '${userId}' to group '${groupId}': ${response.status} ${response.statusText} - ${JSON.stringify(errorData, null, 2)}`)
	}
}

export const removeMemberFromGroup = async (groupId: string, userId: string): Promise<void> => {
	const accessToken = await defaultAzureCredential.getToken(graphScope)
	const response = await fetch(`https://graph.microsoft.com/v1.0/groups/${groupId}/members/${userId}/$ref`, {
		method: "DELETE",
		headers: {
			Authorization: `Bearer ${accessToken?.token}`
		}
	})

	if (!response.ok) {
		const errorData = await response.json()
		throw new Error(`Failed to remove user '${userId}' from group '${groupId}': ${response.status} ${response.statusText} - ${JSON.stringify(errorData, null, 2)}`)
	}
}
