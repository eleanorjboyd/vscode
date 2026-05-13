/*---------------------------------------------------------------------------------------------
 *  Copyright (c) Microsoft Corporation. All rights reserved.
 *  Licensed under the MIT License. See License.txt in the project root for license information.
 *--------------------------------------------------------------------------------------------*/

import assert from 'assert';
import { timeout } from '../../../../../base/common/async.js';
import { Emitter, Event } from '../../../../../base/common/event.js';
import { DisposableStore } from '../../../../../base/common/lifecycle.js';
import { observableValue } from '../../../../../base/common/observable.js';
import { URI } from '../../../../../base/common/uri.js';
import { Selection } from '../../../../../editor/common/core/selection.js';
import { ITextModel } from '../../../../../editor/common/model.js';
import { createTextModel } from '../../../../../editor/test/common/testTextModel.js';
import { ITestCodeEditor, instantiateTestCodeEditor } from '../../../../../editor/test/browser/testCodeEditor.js';
import { TestConfigurationService } from '../../../../../platform/configuration/test/common/testConfigurationService.js';
import { IConfigurationChangeEvent, IConfigurationService } from '../../../../../platform/configuration/common/configuration.js';
import { TestInstantiationService } from '../../../../../platform/instantiation/test/common/instantiationServiceMock.js';
import { ensureNoDisposablesAreLeakedInTestSuite } from '../../../../../base/test/common/utils.js';
import { mock } from '../../../../../base/test/common/mock.js';
import { workbenchInstantiationService } from '../../../../test/browser/workbenchTestServices.js';
import { InlineChatConfigKeys } from '../../common/inlineChat.js';
import { IChatSendRequestOptions, IChatService } from '../../../chat/common/chatService/chatService.js';
import { IInlineChatSession2, IInlineChatSessionService } from '../../browser/inlineChatSessionService.js';
import { InlineChatController } from '../../browser/inlineChatController.js';
import { ChatAgentLocation, ChatModeKind } from '../../../chat/common/constants.js';
import { ILanguageModelChatMetadata, ILanguageModelsService } from '../../../chat/common/languageModels.js';
import { IChatAgentData } from '../../../chat/common/participants/chatAgents.js';
import { IChatModel, IChatResponseModel } from '../../../chat/common/model/chatModel.js';
import { Range } from '../../../../../editor/common/core/range.js';
import { IChatEditingService, IChatEditingSession, IModifiedFileEntry } from '../../../chat/common/editing/chatEditingService.js';
import { Position } from '../../../../../editor/common/core/position.js';
import { CursorChangeReason } from '../../../../../editor/common/cursorEvents.js';
import { CursorState } from '../../../../../editor/common/cursorCommon.js';
import { IUserInteractionService, MockUserInteractionService } from '../../../../../platform/userInteraction/browser/userInteractionService.js';
import { INotebookEditorService } from '../../../notebook/browser/services/notebookEditorService.js';
import { runWithFakedTimers } from '../../../../../base/test/common/timeTravelScheduler.js';
import { IMarkerDecorationsService } from '../../../../../editor/common/services/markerDecorations.js';
import { IMarker, MarkerSeverity } from '../../../../../platform/markers/common/markers.js';
import { IChatRequestVariableEntry } from '../../../chat/common/attachments/chatVariableEntries.js';

suite('InlineChatController - Request Parity', () => {

	const store = new DisposableStore();
	let editor: ITestCodeEditor;
	let model: ITextModel;
	let instantiationService: TestInstantiationService;
	let configurationService: TestConfigurationService;

	/** Captured sendRequest calls: [sessionResource, message, options] */
	let sendRequestCalls: { sessionResource: URI; message: string; options?: IChatSendRequestOptions }[];
	/** Emitter to signal session dispose */
	let sessionDisposedEmitter: Emitter<void>;

	const testModelId = 'test-model-id';
	const testModelQualifiedName = 'Test Model (TestVendor)';
	const testSessionResource = URI.parse('chat-session:test-session');

	setup(() => {
		sendRequestCalls = [];
		sessionDisposedEmitter = store.add(new Emitter<void>());

		instantiationService = workbenchInstantiationService({
			configurationService: () => new TestConfigurationService({
				[InlineChatConfigKeys.RenderMode]: 'hover',
			}),
		}, store);

		configurationService = instantiationService.get(IConfigurationService) as TestConfigurationService;

		// Mock IUserInteractionService — needed for InlineChatInputWidget's internal code editor
		instantiationService.stub(IUserInteractionService, new MockUserInteractionService());

		// Mock INotebookEditorService
		instantiationService.stub(INotebookEditorService, new class extends mock<INotebookEditorService>() {
			override getNotebookForPossibleCell() { return undefined; }
		});

		// Mock IChatService — capture sendRequest calls
		instantiationService.stub(IChatService, new class extends mock<IChatService>() {
			override async sendRequest(sessionResource: URI, message: string, options?: IChatSendRequestOptions) {
				sendRequestCalls.push({ sessionResource, message, options });
				return { kind: 'sent' as const, data: { agent: {} as Partial<IChatAgentData> as IChatAgentData, responseCreatedPromise: Promise.resolve({} as Partial<IChatResponseModel> as IChatResponseModel), responseCompletePromise: Promise.resolve() } };
			}
			override async cancelCurrentRequestForSession() { }
		});

		// Mock ILanguageModelsService
		const testMetadata: ILanguageModelChatMetadata = {
			vendor: 'TestVendor',
			name: 'Test Model',
			family: 'test',
			version: '1',
			id: testModelId,
			maxInputTokens: 1000,
			maxOutputTokens: 1000,
			auth: undefined,
			capabilities: {},
			isDefaultForLocation: { [ChatAgentLocation.EditorInline]: true },
			targetEntitlements: [],
		} as Partial<ILanguageModelChatMetadata> as ILanguageModelChatMetadata;

		instantiationService.stub(ILanguageModelsService, new class extends mock<ILanguageModelsService>() {
			override getLanguageModelIds() { return [testModelId]; }
			override lookupLanguageModel(id: string) { return id === testModelId ? testMetadata : undefined; }
			override lookupLanguageModelByQualifiedName(name: string) {
				if (name === testModelQualifiedName) {
					return { metadata: testMetadata, identifier: testModelId };
				}
				return undefined;
			}
			override async selectLanguageModels() { return [testModelId]; }
		});

		// Mock IChatEditingService
		instantiationService.stub(IChatEditingService, new class extends mock<IChatEditingService>() {
			override readonly editingSessionsObs = observableValue('sessions', []);
		});

		// Mock IInlineChatSessionService
		const onDidChangeSessionsEmitter = store.add(new Emitter<any>());
		const sessionStateObs = observableValue<undefined>('terminationState', undefined);
		const entriesObs = observableValue<readonly IModifiedFileEntry[]>('entries', []);

		instantiationService.stub(IInlineChatSessionService, new class extends mock<IInlineChatSessionService>() {
			override readonly onWillStartSession = Event.None;
			override readonly onDidChangeSessions = onDidChangeSessionsEmitter.event;
			override getSessionByTextModel() { return undefined; }
			override getSessionBySessionUri() { return undefined; }
			override createSession(_editor: any): IInlineChatSession2 {
				const session: IInlineChatSession2 = {
					initialPosition: new Position(1, 1),
					initialSelection: _editor.getSelection() ?? new Selection(1, 1, 1, 6),
					uri: _editor.getModel()!.uri,
					chatModel: {
						sessionResource: testSessionResource,
						initialLocation: ChatAgentLocation.EditorInline,
						hasRequests: false,
						inputModel: { state: observableValue('state', undefined), setState: () => { }, clearState: () => { }, toJSON: () => ({}) },
						getRequests: () => [],
						lastRequestObs: observableValue('lastReq', undefined),
						onDidChange: Event.None,
					} as unknown as IChatModel,
					editingSession: {
						onDidDispose: sessionDisposedEmitter.event,
						entries: entriesObs,
						readEntry: () => undefined,
						getEntry: () => undefined,
						accept: async () => { },
						reject: async () => { },
						dispose: () => { },
					} as Partial<IChatEditingSession> as IChatEditingSession,
					terminationState: sessionStateObs,
					setTerminationState: () => { },
					dispose: () => {
						onDidChangeSessionsEmitter.fire(undefined);
					},
				};
				onDidChangeSessionsEmitter.fire(undefined);
				return session;
			}
		});

		model = store.add(createTextModel('hello world\nfoo bar\nbaz qux'));
		editor = store.add(instantiateTestCodeEditor(instantiationService, model));
	});

	teardown(() => {
		store.clear();
	});

	ensureNoDisposablesAreLeakedInTestSuite();

	function setExplicitSelection(sel: Selection): void {
		editor.getViewModel()!.setCursorStates(
			'test',
			CursorChangeReason.Explicit,
			[CursorState.fromModelSelection(sel)]
		);
	}

	test('hover mode sendRequest has correct location and locationData', () => runWithFakedTimers({ useFakeTimers: true }, async () => {
		setExplicitSelection(new Selection(1, 1, 1, 6));

		const controller = store.add(instantiationService.createInstance(InlineChatController, editor));

		const runPromise = controller.run({ message: 'test message', autoSend: true });
		await timeout(0);

		// Settle the session so run() can return
		sessionDisposedEmitter.fire();
		await runPromise;

		assert.strictEqual(sendRequestCalls.length, 1, 'should have exactly one sendRequest call');
		const call = sendRequestCalls[0];

		// Verify session resource
		assert.ok(call.sessionResource.toString() === testSessionResource.toString());

		// Verify message
		assert.strictEqual(call.message, 'test message');

		// Verify location
		assert.strictEqual(call.options?.location, ChatAgentLocation.EditorInline);

		// Verify locationData
		const locData = call.options?.locationData;
		assert.ok(locData);
		assert.strictEqual(locData.type, ChatAgentLocation.EditorInline);
		if (locData.type === ChatAgentLocation.EditorInline) {
			assert.ok(locData.document.toString() === model.uri.toString());
			assert.deepStrictEqual(Selection.liftSelection(locData.selection), new Selection(1, 1, 1, 6));
		}

		// Verify model selection
		assert.strictEqual(call.options?.userSelectedModelId, testModelId);

		// Verify modeInfo
		assert.strictEqual(call.options?.modeInfo?.kind, ChatModeKind.Ask);
		assert.strictEqual(call.options?.modeInfo?.modeId, 'ask');
		assert.strictEqual(call.options?.modeInfo?.isBuiltin, true);
	}));

	test('hover mode sendRequest locationData matches what zone widget resolveData would produce', () => runWithFakedTimers({ useFakeTimers: true }, async () => {
		setExplicitSelection(new Selection(2, 1, 2, 4));

		const controller = store.add(instantiationService.createInstance(InlineChatController, editor));

		const runPromise = controller.run({ message: 'edit code', autoSend: true });
		await timeout(0);
		sessionDisposedEmitter.fire();
		await runPromise;

		assert.strictEqual(sendRequestCalls.length, 1);
		const locData = sendRequestCalls[0].options?.locationData;
		assert.ok(locData);

		// The zone widget's resolveData builds the same shape:
		// { type: ChatAgentLocation.EditorInline, id: getEditorId(editor, model), selection, document, wholeRange }
		if (locData.type === ChatAgentLocation.EditorInline) {
			// id should be `${editorId},${modelId}`
			assert.ok(typeof locData.id === 'string');
			assert.ok(locData.id.length > 0);
			// document should match the editor's model URI
			assert.ok(locData.document.toString() === model.uri.toString());
			// selection should match what we set
			assert.deepStrictEqual(Selection.liftSelection(locData.selection), new Selection(2, 1, 2, 4));
			// wholeRange should equal the selection (same as zone widget behavior)
			assert.deepStrictEqual(Range.lift(locData.wholeRange), new Range(2, 1, 2, 4));
		} else {
			assert.fail('Expected EditorInline location data');
		}
	}));

	test('hover mode resolves model via defaultModel setting', () => runWithFakedTimers({ useFakeTimers: true }, async () => {
		// Reset _userSelectedModel static
		// @ts-ignore accessing private static for test reset
		InlineChatController._userSelectedModel = undefined;

		// Set a default model config
		configurationService.setUserConfiguration(InlineChatConfigKeys.DefaultModel, testModelQualifiedName);
		configurationService.onDidChangeConfigurationEmitter.fire(new class extends mock<IConfigurationChangeEvent>() {
			override affectsConfiguration() { return true; }
		});

		setExplicitSelection(new Selection(1, 1, 1, 6));
		const controller = store.add(instantiationService.createInstance(InlineChatController, editor));

		const runPromise = controller.run({ message: 'hello', autoSend: true });
		await timeout(0);
		sessionDisposedEmitter.fire();
		await runPromise;

		assert.strictEqual(sendRequestCalls.length, 1);
		assert.strictEqual(sendRequestCalls[0].options?.userSelectedModelId, testModelId);
	}));

	test('hover mode does not send request when autoSend is false', () => runWithFakedTimers({ useFakeTimers: true }, async () => {
		setExplicitSelection(new Selection(1, 1, 1, 6));
		const controller = store.add(instantiationService.createInstance(InlineChatController, editor));

		const runPromise = controller.run({ message: 'hello', autoSend: false });
		await timeout(0);
		sessionDisposedEmitter.fire();
		await runPromise;

		assert.strictEqual(sendRequestCalls.length, 0, 'should not call sendRequest when autoSend is false');
	}));

	test('hover mode does not send request when message is missing', () => runWithFakedTimers({ useFakeTimers: true }, async () => {
		setExplicitSelection(new Selection(1, 1, 1, 6));
		const controller = store.add(instantiationService.createInstance(InlineChatController, editor));

		const runPromise = controller.run({ autoSend: true });
		await timeout(0);
		sessionDisposedEmitter.fire();
		await runPromise;

		assert.strictEqual(sendRequestCalls.length, 0, 'should not call sendRequest when message is missing');
	}));
});

suite('InlineChatController - Zone Mode Diagnostic Attach', () => {

	const store = new DisposableStore();
	let editor: ITestCodeEditor;
	let model: ITextModel;
	let instantiationService: TestInstantiationService;
	let sessionDisposedEmitter: Emitter<void>;

	// Captures from the stubbed zone widget
	let addContextCalls: IChatRequestVariableEntry[][];
	let setInputCalls: string[];
	let inputSetValueCalls: string[];
	let acceptInputCount: number;

	// Markers returned by the stubbed IMarkerDecorationsService
	let liveMarkers: [Range, IMarker][];

	const testModelId = 'test-model-id';
	const testSessionResource = URI.parse('chat-session:test-session-zone');

	function makeMarker(line: number, message: string): [Range, IMarker] {
		const range = new Range(line, 1, line, 10);
		const marker: IMarker = {
			owner: 'test-owner',
			resource: URI.parse('inmemory://model/1'),
			severity: MarkerSeverity.Warning,
			message,
			startLineNumber: line,
			startColumn: 1,
			endLineNumber: line,
			endColumn: 10,
		};
		return [range, marker];
	}

	function installFakeZone(controller: InlineChatController): void {
		const fakeChatWidget = {
			setModel: () => { },
			setInputPlaceholder: () => { },
			setInput: (msg: string) => { setInputCalls.push(msg); },
			acceptInput: async () => { acceptInputCount++; },
			attachmentModel: {
				addContext: (...entries: IChatRequestVariableEntry[]) => { addContextCalls.push(entries); },
				addFile: async () => { },
			},
			input: {
				setValue: (msg: string) => { inputSetValueCalls.push(msg); },
				setCurrentLanguageModel: () => { },
				switchModelByQualifiedName: () => false,
				selectedLanguageModel: observableValue('selectedLM', undefined),
				renderAttachedContext: () => { },
			},
			inputEditor: {
				setSelection: () => { },
			},
		};
		const fakeInlineChatWidget = {
			chatWidget: fakeChatWidget,
			focus: () => { },
			updateInfo: () => { },
			domNode: {
				classList: {
					toggle: () => { },
				},
			},
		};
		const fakeZoneValue = {
			widget: fakeInlineChatWidget,
			position: undefined,
			show: () => { },
			reveal: () => { },
			hide: () => { },
			updatePositionAndHeight: () => { },
		};
		// _zone is a private readonly Lazy<InlineChatZoneWidget>; replace it with a
		// shape-compatible stub so _runZone's zone-widget accesses are intercepted.
		Object.defineProperty(controller, '_zone', { value: { value: fakeZoneValue, rawValue: fakeZoneValue } });
	}

	setup(() => {
		addContextCalls = [];
		setInputCalls = [];
		inputSetValueCalls = [];
		acceptInputCount = 0;
		liveMarkers = [];
		sessionDisposedEmitter = store.add(new Emitter<void>());

		instantiationService = workbenchInstantiationService({
			// Default RenderMode is 'zone'; set explicitly for clarity.
			configurationService: () => new TestConfigurationService({
				[InlineChatConfigKeys.RenderMode]: 'zone',
			}),
		}, store);

		instantiationService.stub(IUserInteractionService, new MockUserInteractionService());

		instantiationService.stub(INotebookEditorService, new class extends mock<INotebookEditorService>() {
			override getNotebookForPossibleCell() { return undefined; }
		});

		instantiationService.stub(IChatService, new class extends mock<IChatService>() {
			override async sendRequest(sessionResource: URI, _message: string, _options?: IChatSendRequestOptions) {
				return { kind: 'sent' as const, data: { agent: {} as Partial<IChatAgentData> as IChatAgentData, responseCreatedPromise: Promise.resolve({} as Partial<IChatResponseModel> as IChatResponseModel), responseCompletePromise: Promise.resolve() } };
			}
			override async cancelCurrentRequestForSession() { }
		});

		const testMetadata: ILanguageModelChatMetadata = {
			vendor: 'TestVendor',
			name: 'Test Model',
			family: 'test',
			version: '1',
			id: testModelId,
			maxInputTokens: 1000,
			maxOutputTokens: 1000,
			auth: undefined,
			capabilities: {},
			isDefaultForLocation: { [ChatAgentLocation.EditorInline]: true },
			targetEntitlements: [],
		} as Partial<ILanguageModelChatMetadata> as ILanguageModelChatMetadata;

		instantiationService.stub(ILanguageModelsService, new class extends mock<ILanguageModelsService>() {
			override getLanguageModelIds() { return [testModelId]; }
			override lookupLanguageModel(id: string) { return id === testModelId ? testMetadata : undefined; }
			override async selectLanguageModels() { return [testModelId]; }
		});

		instantiationService.stub(IChatEditingService, new class extends mock<IChatEditingService>() {
			override readonly editingSessionsObs = observableValue('sessions', []);
		});

		// Marker service returns whatever the test put into liveMarkers
		instantiationService.stub(IMarkerDecorationsService, new class extends mock<IMarkerDecorationsService>() {
			override readonly onDidChangeMarker = Event.None;
			override getLiveMarkers(_uri: URI): [Range, IMarker][] {
				return liveMarkers;
			}
		});

		const onDidChangeSessionsEmitter = store.add(new Emitter<any>());
		const sessionStateObs = observableValue<undefined>('terminationState', undefined);
		const entriesObs = observableValue<readonly IModifiedFileEntry[]>('entries', []);

		instantiationService.stub(IInlineChatSessionService, new class extends mock<IInlineChatSessionService>() {
			override readonly onWillStartSession = Event.None;
			override readonly onDidChangeSessions = onDidChangeSessionsEmitter.event;
			override getSessionByTextModel() { return undefined; }
			override getSessionBySessionUri() { return undefined; }
			override createSession(_editor: any): IInlineChatSession2 {
				const session: IInlineChatSession2 = {
					initialPosition: new Position(1, 1),
					initialSelection: _editor.getSelection() ?? new Selection(1, 1, 1, 6),
					uri: _editor.getModel()!.uri,
					chatModel: {
						sessionResource: testSessionResource,
						initialLocation: ChatAgentLocation.EditorInline,
						hasRequests: false,
						inputModel: { state: observableValue('state', undefined), setState: () => { }, clearState: () => { }, toJSON: () => ({}) },
						getRequests: () => [],
						lastRequestObs: observableValue('lastReq', undefined),
						onDidChange: Event.None,
					} as unknown as IChatModel,
					editingSession: {
						onDidDispose: sessionDisposedEmitter.event,
						entries: entriesObs,
						readEntry: () => undefined,
						getEntry: () => undefined,
						accept: async () => { },
						reject: async () => { },
						dispose: () => { },
					} as Partial<IChatEditingSession> as IChatEditingSession,
					terminationState: sessionStateObs,
					setTerminationState: () => { },
					dispose: () => {
						onDidChangeSessionsEmitter.fire(undefined);
					},
				};
				onDidChangeSessionsEmitter.fire(undefined);
				return session;
			}
		});

		model = store.add(createTextModel('hello world\nfoo bar\nbaz qux\nlast line'));
		editor = store.add(instantiateTestCodeEditor(instantiationService, model));
	});

	teardown(() => {
		store.clear();
	});

	ensureNoDisposablesAreLeakedInTestSuite();

	function setExplicitSelection(sel: Selection): void {
		editor.getViewModel()!.setCursorStates(
			'test',
			CursorChangeReason.Explicit,
			[CursorState.fromModelSelection(sel)]
		);
	}

	test('caller-supplied initialSelection is applied BEFORE diagnostics are collected', () => runWithFakedTimers({ useFakeTimers: true }, async () => {
		// Editor cursor starts at line 1.
		setExplicitSelection(new Selection(1, 1, 1, 1));

		// Two diagnostics: one at line 1 (where the cursor currently is), and one at
		// line 3 (where the caller wants the inline chat to operate).
		liveMarkers = [
			makeMarker(1, 'wrong-marker-line-1'),
			makeMarker(3, 'correct-marker-line-3'),
		];

		const controller = store.add(instantiationService.createInstance(InlineChatController, editor));
		installFakeZone(controller);

		const runPromise = controller.run({
			initialSelection: new Selection(3, 1, 3, 5),
			message: '/fix superfluous-parens',
			autoSend: true,
			attachDiagnostics: true,
		});
		await timeout(0);
		sessionDisposedEmitter.fire();
		await runPromise;

		// Exactly one diagnostic should have been attached, and it must be the
		// one at the caller-supplied selection (line 3), not the one at the
		// editor's pre-invocation cursor (line 1).
		assert.strictEqual(addContextCalls.length, 1, 'addContext should be called exactly once');
		const attached = addContextCalls[0];
		assert.strictEqual(attached.length, 1, 'exactly one diagnostic should be attached');
		const attachedEntry = attached[0];
		if (attachedEntry.kind !== 'diagnostic') {
			assert.fail(`Expected a diagnostic entry but got ${attachedEntry.kind}`);
		}
		assert.strictEqual(attachedEntry.problemMessage, 'correct-marker-line-3');
	}));

	test('caller-supplied initialRange is applied BEFORE diagnostics are collected', () => runWithFakedTimers({ useFakeTimers: true }, async () => {
		setExplicitSelection(new Selection(1, 1, 1, 1));
		liveMarkers = [
			makeMarker(1, 'wrong-marker-line-1'),
			makeMarker(3, 'correct-marker-line-3'),
		];

		const controller = store.add(instantiationService.createInstance(InlineChatController, editor));
		installFakeZone(controller);

		const runPromise = controller.run({
			initialRange: new Range(3, 1, 3, 5),
			message: '/fix superfluous-parens',
			autoSend: true,
			attachDiagnostics: true,
		});
		await timeout(0);
		sessionDisposedEmitter.fire();
		await runPromise;

		assert.strictEqual(addContextCalls.length, 1, 'addContext should be called exactly once');
		const attached = addContextCalls[0];
		assert.strictEqual(attached.length, 1, 'exactly one diagnostic should be attached');
		const attachedEntry = attached[0];
		if (attachedEntry.kind !== 'diagnostic') {
			assert.fail(`Expected a diagnostic entry but got ${attachedEntry.kind}`);
		}
		assert.strictEqual(attachedEntry.problemMessage, 'correct-marker-line-3');
	}));

	test('caller-supplied message is not clobbered when a diagnostic is attached', () => runWithFakedTimers({ useFakeTimers: true }, async () => {
		setExplicitSelection(new Selection(1, 1, 1, 1));
		liveMarkers = [makeMarker(3, 'some-diagnostic')];

		const controller = store.add(instantiationService.createInstance(InlineChatController, editor));
		installFakeZone(controller);

		const callerMessage = '/fix Unnecessary parens after \'if\' keyword';
		const runPromise = controller.run({
			initialSelection: new Selection(3, 1, 3, 5),
			message: callerMessage,
			autoSend: true,
			attachDiagnostics: true,
		});
		await timeout(0);
		sessionDisposedEmitter.fire();
		await runPromise;

		// A diagnostic was attached, but the caller's message must survive — the
		// default "Fix the attached problem" must NOT replace it.
		assert.strictEqual(addContextCalls.length, 1, 'a diagnostic should still be attached');
		assert.ok(
			!inputSetValueCalls.includes('Fix the attached problem'),
			`input.setValue should not have been called with the default message; got: ${JSON.stringify(inputSetValueCalls)}`
		);
		assert.deepStrictEqual(setInputCalls, [callerMessage], 'caller message should be passed to setInput verbatim');
		assert.strictEqual(acceptInputCount, 1, 'autoSend should have triggered acceptInput exactly once');
	}));

	test('default "Fix the attached problem" message is used when caller did not supply one', () => runWithFakedTimers({ useFakeTimers: true }, async () => {
		setExplicitSelection(new Selection(1, 1, 1, 1));
		liveMarkers = [makeMarker(3, 'some-diagnostic')];

		const controller = store.add(instantiationService.createInstance(InlineChatController, editor));
		installFakeZone(controller);

		const runPromise = controller.run({
			initialSelection: new Selection(3, 1, 3, 5),
			autoSend: true,
			attachDiagnostics: true,
		});
		await timeout(0);
		sessionDisposedEmitter.fire();
		await runPromise;

		// No caller message → the default "Fix the attached problem" should be
		// set on the input AND submitted via acceptInput.
		assert.strictEqual(addContextCalls.length, 1, 'a diagnostic should be attached');
		assert.deepStrictEqual(inputSetValueCalls, ['Fix the attached problem']);
		assert.deepStrictEqual(setInputCalls, ['Fix the attached problem']);
		assert.strictEqual(acceptInputCount, 1);
	}));
});
