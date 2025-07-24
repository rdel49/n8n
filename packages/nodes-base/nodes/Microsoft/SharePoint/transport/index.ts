import type {
	IDataObject,
	IExecuteFunctions,
	IExecuteSingleFunctions,
	IHttpRequestMethods,
	IHttpRequestOptions,
	ILoadOptionsFunctions,
} from 'n8n-workflow';

export async function microsoftSharePointApiRequest(
	this: IExecuteFunctions | IExecuteSingleFunctions | ILoadOptionsFunctions,
	method: IHttpRequestMethods,
	endpoint: string,
	body: IDataObject | Buffer = {},
	qs?: IDataObject,
	headers?: IDataObject,
	url?: string,
): Promise<any> {
	const credentials: { subdomain?: string; customDomain?: string } = await this.getCredentials(
		'microsoftSharePointOAuth2Api',
	);

	const domain = credentials.customDomain?.trim()
		? credentials.customDomain.trim()
		: `${credentials.subdomain}.sharepoint.com`;

	const resolvedUrl = url ?? `https://${domain}/_api/v2.0${endpoint}`;

	const options: IHttpRequestOptions = {
		method,
		url: resolvedUrl,
		json: true,
		headers,
		body,
		qs,
	};

	console.log('SharePoint request URL:', resolvedUrl);

	return await this.helpers.httpRequestWithAuthentication.call(
		this,
		'microsoftSharePointOAuth2Api',
		options,
	);
}
