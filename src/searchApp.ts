import {
  MessageExtensionTokenResponse,
  OnBehalfOfUserCredential,
  handleMessageExtensionQueryWithSSO
} from "@microsoft/teamsfx";
import {
  TeamsActivityHandler,
  CardFactory,
  TurnContext,
  MessagingExtensionQuery,
  CardImage,
  AttachmentLayoutTypes,
  MessagingExtensionResponse,
} from "botbuilder";
import {
  Client,
  ResponseType
} from "@microsoft/microsoft-graph-client";
import { TokenCredentialAuthenticationProvider } from "@microsoft/microsoft-graph-client/authProviders/azureTokenCredentials";
import config from "./config";

const listFields = [
  "fields/Title",
  "fields/RetailCategory",
  "fields/Specguide",
  "fields/PhotoSubmission",
  "fields/CustomerRating",
  "fields/ReleaseDate"
];

export class SearchApp extends TeamsActivityHandler {
  constructor() {
    super();
  }

  // Search.
  public async handleTeamsMessagingExtensionQuery(
    context: TurnContext,
    query: MessagingExtensionQuery
  ): Promise<any> {
    return await handleMessageExtensionQueryWithSSO(
      context,
      config.authConfig,
      config.initiateLoginEndpoint,
      "Sites.Read.All",
      async (token: MessageExtensionTokenResponse) => {
        const credential = new OnBehalfOfUserCredential(token.ssoToken, config.authConfig);
        const authProvider = new TokenCredentialAuthenticationProvider(credential, { scopes: ["Sites.Read.All"] });
        const graphClient = Client.initWithMiddleware({ authProvider: authProvider });

        const { sharepointIds } = await graphClient.api(`/sites/${config.spoHostname}:/${config.spoSiteUrl}`).select("sharepointIds").get();
        const { value: items } = await graphClient.api(`/sites/${sharepointIds.siteId}/lists/Products/items?expand=fields&select=${listFields.join(",")}&$filter=startswith(fields/Title,'${query.parameters[0].value}')`).get();
        const { value: drives } = await graphClient.api(`sites/${sharepointIds.siteId}/drives`).select(["id", "name"]).get();
        const drive = drives.find(drive => drive.name === "Product Imagery");

        const attachments = [];
        await Promise.all(items.map(async (item) => {
          const { PhotoSubmission: photoUrl, Title, RetailCategory } = item.fields;
          const fileName = photoUrl.split("/").reverse()[0];
          const driveItem = await graphClient.api(`sites/${sharepointIds.siteId}/drives/${drive.id}/root:/${fileName}`).get();
          const content = await graphClient.api(`sites/${sharepointIds.siteId}/drives/${drive.id}/items/${driveItem.id}/content`).responseType(ResponseType.ARRAYBUFFER).get();
          const cardImages: CardImage[] = [{ url: `data:${driveItem.file.mimeType};base64,${Buffer.from(content).toString('base64')}`, alt: Title }]
          const card = CardFactory.thumbnailCard(Title, RetailCategory, cardImages);
          attachments.push(card);
        }));

        return {
          composeExtension: {
            type: "result",
            attachmentLayout: AttachmentLayoutTypes.List,
            attachments,
          },
        } as MessagingExtensionResponse;
      }
    );
  }
}