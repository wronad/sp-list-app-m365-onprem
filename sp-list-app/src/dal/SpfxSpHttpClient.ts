import { SPHttpClient } from "@microsoft/sp-http";
import { IHttpClient, IHttpClientResponse } from "./IHttpClient";

const buildHeader = (options?: RequestInit): any => {
  return {
    ...options?.headers,
    Accept: "application/json",
    "Content-Type": "application/json",
  };
};

export class SpfxSpHttpClient implements IHttpClient {
  constructor(protected httpClient: SPHttpClient) {}
  public get(url: string, options?: RequestInit): Promise<IHttpClientResponse> {
    return this.httpClient.get(url, SPHttpClient.configurations.v1, {
      ...options,
      headers: buildHeader(options),
    });
  }
  public post(
    url: string,
    options?: RequestInit
  ): Promise<IHttpClientResponse> {
    return this.httpClient.post(url, SPHttpClient.configurations.v1, {
      ...options,
      headers: buildHeader(options),
    });
  }
  public patch(
    url: string,
    options?: RequestInit
  ): Promise<IHttpClientResponse> {
    return this.httpClient.fetch(url, SPHttpClient.configurations.v1, {
      ...options,
      headers: buildHeader(options),
      method: "PATCH",
    });
  }
  public put(url: string, options?: RequestInit): Promise<IHttpClientResponse> {
    return this.httpClient.fetch(url, SPHttpClient.configurations.v1, {
      ...options,
      headers: buildHeader(options),
      method: "PUT",
    });
  }
  public delete(url: string): Promise<IHttpClientResponse> {
    return this.httpClient.fetch(url, SPHttpClient.configurations.v1, {
      method: "DELETE",
    });
  }
}
