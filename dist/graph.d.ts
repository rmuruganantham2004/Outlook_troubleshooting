import "isomorphic-fetch";
import { Client } from "@microsoft/microsoft-graph-client";
export declare function getGraphClient(): Promise<Client>;
