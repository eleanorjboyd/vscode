/*---------------------------------------------------------------------------------------------
 *  Copyright (c) Microsoft Corporation. All rights reserved.
 *  Licensed under the MIT License. See License.txt in the project root for license information.
 *--------------------------------------------------------------------------------------------*/

import assert from 'assert';
import { DeferredPromise, timeout } from '../../../../base/common/async.js';
import { DisposableStore } from '../../../../base/common/lifecycle.js';
import { constObservable } from '../../../../base/common/observable.js';
import { URI } from '../../../../base/common/uri.js';
import { IChannelServer, IServerChannel } from '../../../../base/parts/ipc/common/ipc.js';
import { ensureNoDisposablesAreLeakedInTestSuite } from '../../../../base/test/common/utils.js';
import { runWithFakedTimers } from '../../../../base/test/common/timeTravelScheduler.js';
import { TestConfigurationService } from '../../../configuration/test/common/testConfigurationService.js';
import { IEnvironmentService } from '../../../environment/common/environment.js';
import { IInstantiationService } from '../../../instantiation/common/instantiation.js';
import { TestInstantiationService } from '../../../instantiation/test/common/instantiationServiceMock.js';
import { NullLogService } from '../../../log/common/log.js';
import { agentsWindowAgentHostClientInfo } from '../../common/agentHostClientInfo.js';
import { IAgentHostEnablementService } from '../../common/agentHostEnablementService.js';
import { AGENT_HOST_CLIENT_PROXY_CHANNEL } from '../../common/agentHostClientProxyChannel.js';
import { AGENT_HOST_CLIENT_BYOK_LM_CHANNEL, AgentHostClientByokLmChannel } from '../../common/agentHostClientByokLmChannel.js';
import { IAgentCreateSessionConfig, IAgentHostManagementService } from '../../common/agentService.js';
import { LocalAgentHostServiceClient, registerAgentHostClientChannels } from '../../electron-browser/localAgentHostService.js';

suite('LocalAgentHostServiceClient', () => {

	ensureNoDisposablesAreLeakedInTestSuite();

	test('bounds extension-bearing createSession when management IPC does not settle', () => {
		return runWithFakedTimers({ useFakeTimers: true, maxTaskCount: 10_000 }, async () => {
			const disposables = new DisposableStore();
			const hanging = new DeferredPromise<URI>();
			const session = URI.parse('ahp-copilotcli://session/hanging');
			let receivedConfig: IAgentCreateSessionConfig | undefined;
			const managementService: IAgentHostManagementService = {
				_serviceBrand: undefined,
				createSessionWithExtensions: config => {
					receivedConfig = config;
					return hanging.p;
				},
				createChatWithExtensions: async () => assert.fail('Unexpected createChatWithExtensions call'),
				shutdown: async () => assert.fail('Unexpected shutdown call'),
				getNetworkDiagnosticsInfo: async () => assert.fail('Unexpected getNetworkDiagnosticsInfo call'),
				getManagedSettingsDiagnostics: async () => assert.fail('Unexpected getManagedSettingsDiagnostics call'),
				diagnosticsFetch: async () => assert.fail('Unexpected diagnosticsFetch call'),
				startWebSocketServer: async () => assert.fail('Unexpected startWebSocketServer call'),
				getInspectInfo: async () => assert.fail('Unexpected getInspectInfo call'),
			};
			const enablementService: IAgentHostEnablementService = {
				_serviceBrand: undefined,
				enabled: constObservable(false),
			};
			const instantiationService = disposables.add(new TestInstantiationService());
			const client = disposables.add(new LocalAgentHostServiceClient(
				agentsWindowAgentHostClientInfo,
				new NullLogService(),
				new TestConfigurationService(),
				{ logsHome: URI.file('/logs') } as IEnvironmentService,
				instantiationService,
				enablementService,
			));
			let trackedSession: URI | undefined;
			let trackedPromise: Promise<unknown> | undefined;
			Object.defineProperty(client, '_protocolClient', {
				value: {
					trackSessionCreate: (resource: URI, promise: Promise<unknown>) => {
						trackedSession = resource;
						trackedPromise = promise;
					},
				},
			});
			Object.defineProperty(client, '_callManagement', {
				value: <T>(callback: (management: IAgentHostManagementService) => Promise<T>) => callback(managementService),
			});

			try {
				const creation = client.createSession({ provider: 'copilotcli', model: { id: 'gpt-4' }, session });
				const rejection = creation.catch(error => error);

				assert.strictEqual(receivedConfig?.session, session);
				assert.strictEqual(trackedSession, session);
				assert.strictEqual(trackedPromise, creation);

				await timeout(31_000);
				const error = await rejection;
				assert.ok(error instanceof Error);
				assert.match(error.message, /session creation for ahp-copilotcli:\/\/session\/hanging timed out after 30000ms/);
			} finally {
				hanging.complete(session);
				disposables.dispose();
			}
		});
	});
});

/**
 * Regression coverage for the renderer reverse-RPC channel registration. The
 * BYOK language-model bridge depends on `IAgentHostByokLmHandler`, registered by
 * the chat contribution of every window that backs BYOK (the main workbench and
 * the Agents app). The registration must still degrade gracefully should a
 * window ever connect without binding the handler — `createInstance` then throws,
 * and that must NOT abort the rest of `_connect` (client completion,
 * action/notification wiring, root-state subscription), or the whole window loses
 * its agent host.
 */
suite('registerAgentHostClientChannels', () => {

	ensureNoDisposablesAreLeakedInTestSuite();

	function fakeChannelServer(): { server: IChannelServer; registered: string[] } {
		const registered: string[] = [];
		const server: IChannelServer = {
			registerChannel: (name: string, _channel: IServerChannel) => { registered.push(name); },
		};
		return { server, registered };
	}

	/**
	 * Minimal {@link IInstantiationService} whose `createInstance` throws for the
	 * BYOK channel when `byokHandlerMissing`, mirroring the strict "UNKNOWN
	 * service agentHostByokLmHandler" failure in windows without the handler.
	 */
	function fakeInstantiationService(byokHandlerMissing: boolean): IInstantiationService {
		return {
			createInstance: (ctor: unknown) => {
				if (ctor === AgentHostClientByokLmChannel && byokHandlerMissing) {
					throw new Error('[createInstance] AgentHostClientByokLmChannel depends on UNKNOWN service agentHostByokLmHandler.');
				}
				return {};
			},
		} as unknown as IInstantiationService;
	}

	test('registers both channels when BYOK is enabled and the handler is available', () => {
		const { server, registered } = fakeChannelServer();
		registerAgentHostClientChannels(server, fakeInstantiationService(false), new NullLogService(), true);
		assert.deepStrictEqual(registered, [AGENT_HOST_CLIENT_PROXY_CHANNEL, AGENT_HOST_CLIENT_BYOK_LM_CHANNEL]);
	});

	test('registers only the proxy channel and does NOT throw when the BYOK handler is missing', () => {
		const { server, registered } = fakeChannelServer();
		// Must not throw: the agent host connection has to come up even if a
		// window connects without the handler and so cannot serve BYOK itself.
		registerAgentHostClientChannels(server, fakeInstantiationService(true), new NullLogService(), true);
		assert.deepStrictEqual(registered, [AGENT_HOST_CLIENT_PROXY_CHANNEL]);
	});

	test('registers only the proxy channel when BYOK is disabled', () => {
		const { server, registered } = fakeChannelServer();
		registerAgentHostClientChannels(server, fakeInstantiationService(false), new NullLogService(), false);
		assert.deepStrictEqual(registered, [AGENT_HOST_CLIENT_PROXY_CHANNEL]);
	});
});
