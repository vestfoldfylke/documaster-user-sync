const ENV_GROUP_MAPPINGS = Object.entries(process.env).filter(([key]) => key.startsWith("SP_ROW"))

export const GROUP_MAPPINGS: Record<string, string> = {}

for (const [_, value] of ENV_GROUP_MAPPINGS) {
	if (typeof value !== "string") {
		throw new Error(`SP_ROW environment variable must be a string, got: ${typeof value}`)
	}
	const keyValue = value.split(";")
	if (keyValue.length !== 2) {
		throw new Error(`Invalid SP_ROW environment variable format: ${value}. Expected format is 'SharePointGroupName;AzureADGroupId'`)
	}
	if (!keyValue[0] || !keyValue[1]) {
		throw new Error(`Invalid SP_ROW environment variable format: ${value}. SharePointGroupName and AzureADGroupId cannot be empty`)
	}
	GROUP_MAPPINGS[keyValue[0]] = keyValue[1]
}

if (!process.env.DOCUMASTER_SOURCE_LIST_SITE_ID || !process.env.DOCUMASTER_SOURCE_LIST_LIST_ID) {
	throw new Error("DOCUMASTER_SOURCE_LIST_SITE_ID and DOCUMASTER_SOURCE_LIST_LIST_ID must be set in environment variables")
}

export const DOCUMASTER_LIST = {
	SITE_ID: process.env.DOCUMASTER_SOURCE_LIST_SITE_ID,
	LIST_ID: process.env.DOCUMASTER_SOURCE_LIST_LIST_ID
}
