import { DefaultAzureCredential } from "@azure/identity"
import type { ListItem, User } from "@microsoft/microsoft-graph-types"

if (!process.env.AZURE_TENANT_ID || !process.env.AZURE_CLIENT_ID || !process.env.AZURE_CLIENT_SECRET) {
	throw new Error("AZURE_TENANT_ID, AZURE_CLIENT_ID, and AZURE_CLIENT_SECRET must be set in environment variables")
}

const defaultAzureCredential = new DefaultAzureCredential({})

const graphScope = "https://graph.microsoft.com/.default"

async function callGraph<T>(method: "GET" | "POST" | "PATCH" | "DELETE", endpoint: string, body?: unknown): Promise<T> {
	const accessToken = await defaultAzureCredential.getToken(graphScope)
	const requestInit: RequestInit = {
		method,
		headers: {
			Authorization: `Bearer ${accessToken?.token}`
		}
	}
	if (body) {
		requestInit.headers = {
			...requestInit.headers,
			"Content-Type": "application/json"
		}
		requestInit.body = JSON.stringify(body)
	}
	const response = await fetch(`https://graph.microsoft.com/v1.0/${endpoint}`, requestInit)

	if (!response.ok) {
		const errorData = await response.json()
		throw new Error(`Failed to fetch endpoint '${endpoint}' : ${response.status} ${response.statusText} - ${JSON.stringify(errorData, null, 2)}`)
	}

	return (await response.json()) as T
}

export const getListItems = async (siteId: string, listId: string): Promise<ListItem[]> => {
	return (await callGraph<{ value: ListItem[] }>("GET", `sites/${siteId}/lists/${listId}/items?expand=fields`)).value
}

export const updateListItemFields = async (siteId: string, listId: string, itemId: string, fields: Record<string, unknown>): Promise<void> => {
	await callGraph("PATCH", `sites/${siteId}/lists/${listId}/items/${itemId}/fields`, fields)
}

export const getGroupMembers = async (groupId: string): Promise<User[]> => {
	return (await callGraph<{ value: User[] }>("GET", `groups/${groupId}/members`)).value
}

export const getUserByUpn = async (email: string): Promise<User | null> => {
	const result = await callGraph<{ value: User[] }>("GET", `users?$filter=mail eq '${email}' or userPrincipalName eq '${email}'`)
	return result.value.length > 0 && result.value[0] !== undefined ? result.value[0] : null
}

export const addMemberToGroup = async (groupId: string, userId: string): Promise<void> => {
	await callGraph("POST", `groups/${groupId}/members/$ref`, {
		"@odata.id": `https://graph.microsoft.com/v1.0/users/${userId}`
	})
}

export const removeMemberFromGroup = async (groupId: string, userId: string): Promise<void> => {
	await callGraph("DELETE", `groups/${groupId}/members/${userId}/$ref`)
}
