import type { Icon, ICredentialType, INodeProperties } from 'n8n-workflow';

export class MicrosoftSharePointOAuth2Api implements ICredentialType {
	name = 'microsoftSharePointOAuth2Api';

	extends = ['microsoftOAuth2Api'];

	icon: Icon = {
		light: 'file:icons/microsoftSharePoint.svg',
		dark: 'file:icons/microsoftSharePoint.svg',
	};

	displayName = 'Microsoft SharePoint OAuth2 API';

	documentationUrl = 'microsoft';

	httpRequestNode = {
		name: 'Microsoft SharePoint',
		docsUrl: 'https://learn.microsoft.com/en-us/sharepoint/dev/apis/sharepoint-rest-graph',
		apiBaseUrlPlaceholder:
			'https://{{ $self.customDomain ? $self.customDomain : `${$self.subdomain}.sharepoint.com` }}/_api/v2.0/',
	};

	properties: INodeProperties[] = [
		{
			displayName: 'Scope',
			name: 'scope',
			type: 'hidden',
			default:
				'=openid offline_access https://{{ $self.customDomain ? $self.customDomain : `${$self.subdomain}.sharepoint.com` }}/.default',
		},
		{
			displayName: 'Subdomain',
			name: 'subdomain',
			type: 'string',
			default: '',
			hint: 'You can extract the subdomain from the URL. For example, in the URL "https://tenant123.sharepoint.com", the subdomain is "tenant123".',
		},
		{
			name: 'customDomain',
			displayName: 'Custom Domain',
			type: 'string',
			default: '',
			hint: 'Optional: Only provide a full subdomain and domain if you need to override the default sharepoint.com domain.',
		},
	];
}
