import type * as GraphTypes from "@microsoft/microsoft-graph-types"

export type DocumasterGraphListItem = GraphTypes.ListItem & {
	fields: {
		Title: string
		Hartilgang: {
			LookupId: number
			LookupValue: string
			Email: string
		}[]
	}
}
