/* Office JavaScript OM library */
/* Version: 15.0.3612 */
/*
	Copyright (c) Microsoft Corporation.  All rights reserved.
*/

if (typeof OSF=="undefined") {
	OSF={};
}
OSF.OUtil=(function () {
	var _uniqueId=-1;
	var _xdmInfoKey='&_xdm_Info=';
	var _xdmCookieNamePrefix='_xdm_';
	var _cacheKeyPrefix='__OSF_XDM.';
	var _fragmentSeparator='#';
	var _loadedScripts={};
	function _random() {
		return Math.floor(100000001 * Math.random()).toString();
	};
	return {
		extend: function OSF_OUtil$extend(child, parent) {
			var F=function () { };
			F.prototype=parent.prototype;
			child.prototype=new F();
			child.prototype.constructor=child;
			child.uber=parent.prototype;
			if (parent.prototype.constructor===Object.prototype.constructor) {
				parent.prototype.constructor=parent;
			}
		},
		setNamespace: function OSF_OUtil$setNamespace(name, parent) {
			if (parent !==undefined && name !==undefined && parent[name]===undefined) {
				parent[name]={};
			}
		},
		unsetNamespace: function OSF_OUtil$unsetNamespace(name, parent) {
			if (parent !==undefined && name !==undefined && parent[name] !==undefined) {
				delete parent[name];
			}
		},
		loadScript: function OSF_OUtil$loadScript(url, callback) {
			if (url && callback) {
				var doc=window.document;
				var _loadedScriptEntry=_loadedScripts[url];
				if(!_loadedScriptEntry) {
					var script=doc.createElement("script");
					script.type="text/javascript";
					_loadedScriptEntry={loaded : false, pendingCallbacks : [callback]};
					_loadedScripts[url]=_loadedScriptEntry;
					var onLoadCallback=function() {
						_loadedScriptEntry.loaded=true;
						var pendingCallbackCount=_loadedScriptEntry.pendingCallbacks.length;
						for(var i=0; i < pendingCallbackCount; i++) {
							var currentCallback=_loadedScriptEntry.pendingCallbacks.shift();
							currentCallback();
						}
					};
					if (script.readyState) {
						script.onreadystatechange=function () {
							if (script.readyState=="loaded" || script.readyState=="complete") {
								script.onreadystatechange=null;
								onLoadCallback();
							}
						};
					} else {
						script.onload=onLoadCallback;
					}
					script.src=url;
					doc.getElementsByTagName("head")[0].appendChild(script);
				} else if(_loadedScriptEntry.loaded) {
					callback();
				} else {
					_loadedScriptEntry.pendingCallbacks.push(callback);
				}
			}
		},
		loadCSS: function OSF_OUtil$loadCSS(url) {
			if (url) {
				var doc=window.document;
				var link=doc.createElement("link");
				link.type="text/css";
				link.rel="stylesheet";
				link.href=url;
				doc.getElementsByTagName("head")[0].appendChild(link);
			}
		},
		parseEnum: function OSF_OUtil$parseEnum(str, enumObject) {
		  var parsed=enumObject[str.trim()];
		  if (typeof (parsed)=='undefined')  {
			Sys.Debug.trace("invalid enumeration string:"+str);
			throw Error.argument("str");
		  }
		  return parsed;
		},
		getUniqueId: function OSF_OUtil$getUniqueId() {
			_uniqueId=_uniqueId+1;
			return _uniqueId.toString();
		},
		formatString: function OSF_OUtil$formatString() {
			var args=arguments;
			var source=args[0];
			return source.replace(/{(\d+)}/gm, function (match, number) {
				var index=parseInt(number, 10)+1;
				return args[index]===undefined ? '{'+number+'}' : args[index];
			});
		},
		generateConversationId: function OSF_OUtil$generateConversationId() {
			return [_random(), _random(), (new Date()).getTime().toString()].join('_');
		},
		getFrameNameAndConversationId: function OSF_OUtil$getFrameNameAndConversationId(cacheKey, frame) {
			var sessionKey=_cacheKeyPrefix+cacheKey;
			var frameName=null;
			var conversationId=null;
			if(window.sessionStorage) {
				var cacheEntry=window.sessionStorage.getItem(sessionKey);
				if(cacheEntry) {
					var parts=cacheEntry.split("|");
					frameName=parts[0];
					conversationId=parts[1];
				}
			}
			if(!frameName) {
				frameName=_xdmCookieNamePrefix+this.generateConversationId();
				conversationId=this.generateConversationId();
				if(window.sessionStorage) {
					window.sessionStorage.setItem(sessionKey, frameName+"|"+conversationId);
				}
			}
			frame.setAttribute("name", frameName);
			return  conversationId;
		},
		addXdmInfoAsHash: function OSF_OUtil$addXdmInfoAsHash(url, xdmInfoValue) {
			url=url.trim() || '';
			var urlParts=url.split(_fragmentSeparator);
			var urlWithoutFragment=urlParts.shift();
			var fragment=urlParts.join(_fragmentSeparator);
			return [urlWithoutFragment, _fragmentSeparator, fragment, _xdmInfoKey, xdmInfoValue].join('');
		},
		parseXdmInfo: function OSF_OUtil$parseXdmInfo() {
			var fragment=window.location.hash;
			var fragmentParts=fragment.split(_xdmInfoKey);
			var xdmInfoValue=fragmentParts.length > 1 ? fragmentParts[fragmentParts.length - 1] : null;
			var cookieNameStart=window.name.indexOf(_xdmCookieNamePrefix);
			if(cookieNameStart > -1) {
				var cookieNameEnd=window.name.indexOf(";", cookieNameStart);
				if (cookieNameEnd==-1){
					cookieNameEnd=window.name.length;
				}
				var cookieName=window.name.substring(cookieNameStart, cookieNameEnd);
				if(xdmInfoValue) {
					document.cookie=encodeURIComponent(cookieName)+"="+encodeURIComponent(xdmInfoValue);
				} else {
					var cookieKey=encodeURIComponent(cookieName)+"=";
					var cookieStart=document.cookie.indexOf(cookieKey);
					if (cookieStart > -1){
						var cookieEnd=document.cookie.indexOf(";", cookieStart);
						if (cookieEnd==-1){
							cookieEnd=document.cookie.length;
						}
						xdmInfoValue=decodeURIComponent(document.cookie.substring(cookieStart+cookieKey.length, cookieEnd));
					}
				}
			}
			return xdmInfoValue;
		},
		getTrailingItem: function OSF_OUtil$getTrailingFunction(list, type) {
			if (list.length > 0) {
				var candidate=list[list.length - 1];
				if (typeof candidate==type)
					return candidate;
			}
			return null;
		},
		checkParamsAndGetCallback: function OSF_OUtil$checkParamsAndGetCallback(suppliedArguments, expectedArguments) {
			var callback=OSF.OUtil.getTrailingItem(suppliedArguments, "function");
			var options=OSF.OUtil.getTrailingItem(suppliedArguments, "object");
			if (options) {
				if (options[Microsoft.Office.WebExtension.OptionalParameters.Callback]) {
					if (callback) {
						throw OSF.OUtil.formatString(Strings.OfficeOM.L_RedundantCallbackSpecification);
					} else {
						callback=options[Microsoft.Office.WebExtension.OptionalParameters.Callback];
						var callbackType=typeof callback;
						if (callbackType !="function") {
							throw OSF.OUtil.formatString(Strings.OfficeOM.L_CallbackNotAFunction, callbackType);
						}
					}
				}
			}
			expectedArguments.push({ name: "options", type: Object, optional: true });
			var e=Function._validateParams(suppliedArguments, expectedArguments, false );
			if (e) throw e;
			return callback;
		},
		validateParamObject: function OSF_OUtil$validateParamObject(params, expectedProperties, callback) {
			var e=Function._validateParams(arguments, [
				{ name: "params", type: Object, mayBeNull: false },
				{ name: "expectedProperties", type: Object, mayBeNull: false },
				{ name: "callback", type: Function, mayBeNull: true }
			]);
			if (e) throw e;
			for(var p in expectedProperties) {
				e=Function._validateParameter(params[p], expectedProperties[p], p);
				if (e) throw e;
			}
		},
		writeProfilerMark: function OSF_OUtil$writeProfilerMark(text)
		{
			if (window.msWriteProfilerMark) window.msWriteProfilerMark(text);
		},
		addEventListener: function OSF_OUtil$addEventListener(element, eventName, listener)
		{
			if (element.attachEvent) {
				element.attachEvent("on"+eventName, listener);
			} else if(element.addEventListener){
				element.addEventListener(eventName, listener, false);
			} else {
				element["on"+eventName]=listener;
			}
		},
		removeEventListener: function OSF_OUtil$removeEventListener(element, eventName, listener)
		{
			if (element.detachEvent) {
				element.detachEvent("on"+eventName, listener);
			} else if(element.removeEventListener){
				element.removeEventListener(eventName, listener, false);
			} else {
				element["on"+eventName]=null;
			}
		}
	};
})();
OSF.OUtil.setNamespace("Microsoft", window);
OSF.OUtil.setNamespace("Office", Microsoft);
OSF.OUtil.setNamespace("Common", Microsoft.Office);
Microsoft.Office.Common.InvokeType={ "async": 0,
									   "sync": 1,
									   "asyncRegisterEvent": 2,
									   "asyncUnregisterEvent": 3,
									   "syncRegisterEvent": 4,
									   "syncUnregisterEvent": 5
									   };
Microsoft.Office.Common.InvokeResultCode={
											 "noError": 0,
											 "errorInRequest": -1,
											 "errorHandlingRequest": -2,
											 "errorInResponse": -3,
											 "errorHandlingResponse": -4,
											 "errorHandlingRequestAccessDenied": -5,
											 "errorHandlingMethodCallTimedout": -6
											};
Microsoft.Office.Common.MessageType={ "request": 0,
										"response": 1
									  };
Microsoft.Office.Common.ActionType={ "invoke": 0,
									   "registerEvent": 1,
									   "unregisterEvent": 2 };
Microsoft.Office.Common.ResponseType={ "forCalling": 0,
										 "forEventing": 1
									  };
Microsoft.Office.Common.MethodObject=function Microsoft_Office_Common_MethodObject(method, invokeType, blockingOthers) {
	this._method=method;
	this._invokeType=invokeType;
	this._blockingOthers=blockingOthers;
};
Microsoft.Office.Common.MethodObject.prototype={
	getMethod: function Microsoft_Office_Common_MethodObject$getMethod() {
		return this._method;
	},
	getInvokeType: function Microsoft_Office_Common_MethodObject$getInvokeType() {
		return this._invokeType;
	},
	getBlockingFlag: function Microsoft_Office_Common_MethodObject$getBlockingFlag() {
		return this._blockingOthers;
	}
};
Microsoft.Office.Common.EventMethodObject=function Microsoft_Office_Common_EventMethodObject(registerMethodObject, unregisterMethodObject) {
	this._registerMethodObject=registerMethodObject;
	this._unregisterMethodObject=unregisterMethodObject;
};
Microsoft.Office.Common.EventMethodObject.prototype={
	getRegisterMethodObject: function Microsoft_Office_Common_EventMethodObject$getRegisterMethodObject() {
		return this._registerMethodObject;
	},
	getUnregisterMethodObject: function Microsoft_Office_Common_EventMethodObject$getUnregisterMethodObject() {
		return this._unregisterMethodObject;
	}
};
Microsoft.Office.Common.ServiceEndPoint=function Microsoft_Office_Common_ServiceEndPoint(serviceEndPointId) {
	var e=Function._validateParams(arguments, [
		  { name: "serviceEndPointId", type: String, mayBeNull: false }
	]);
	if (e) throw e;
	this._methodObjectList={};
	this._eventHandlerProxyList={};
	this._Id=serviceEndPointId;
	this._conversations={};
	this._policyManager=null;
};
Microsoft.Office.Common.ServiceEndPoint.prototype={
	registerMethod: function Microsoft_Office_Common_ServiceEndPoint$registerMethod(methodName, method, invokeType, blockingOthers) {
		var e=Function._validateParams(arguments, [
			{ name: "methodName", type: String, mayBeNull: false },
			{ name: "method", type: Function, mayBeNull: false },
			{ name: "invokeType", type: Number, mayBeNull: false },
			{ name: "blockingOthers", type: Boolean, mayBeNull: false }
		]);
		if (e) throw e;
		if (invokeType !==Microsoft.Office.Common.InvokeType.async
			&& invokeType !==Microsoft.Office.Common.InvokeType.sync){
			throw Error.argument("invokeType");
		}
		var methodObject=new Microsoft.Office.Common.MethodObject(method,
																	invokeType,
																	blockingOthers);
		this._methodObjectList[methodName]=methodObject;
	},
	unregisterMethod: function Microsoft_Office_Common_ServiceEndPoint$unregisterMethod(methodName) {
		var e=Function._validateParams(arguments, [
			{ name: "methodName", type: String, mayBeNull: false }
		]);
		if (e) throw e;
		delete this._methodObjectList[methodName];
	},
	registerEvent: function Microsoft_Office_Common_ServiceEndPoint$registerEvent(eventName, registerMethod, unregisterMethod) {
		var e=Function._validateParams(arguments, [
			{ name: "eventName", type: String, mayBeNull: false },
			{ name: "registerMethod", type: Function, mayBeNull: false },
			{ name: "unregisterMethod", type: Function, mayBeNull: false }
		]);
		if (e) throw e;
		var methodObject=new Microsoft.Office.Common.EventMethodObject (
																		  new Microsoft.Office.Common.MethodObject(registerMethod,
																												   Microsoft.Office.Common.InvokeType.syncRegisterEvent,
																												   false),
																		  new Microsoft.Office.Common.MethodObject(unregisterMethod,
																												   Microsoft.Office.Common.InvokeType.syncUnregisterEvent,
																												   false)
																												   );
		this._methodObjectList[eventName]=methodObject;
	},
	registerEventEx: function Microsoft_Office_Common_ServiceEndPoint$registerEventEx(eventName, registerMethod, registerMethodInvokeType, unregisterMethod, unregisterMethodInvokeType) {
		var e=Function._validateParams(arguments, [
			{ name: "eventName", type: String, mayBeNull: false },
			{ name: "registerMethod", type: Function, mayBeNull: false },
			{ name: "registerMethodInvokeType", type: Number, mayBeNull: false },
			{ name: "unregisterMethod", type: Function, mayBeNull: false },
			{ name: "unregisterMethodInvokeType", type: Number, mayBeNull: false }
		]);
		if (e) throw e;
		var methodObject=new Microsoft.Office.Common.EventMethodObject (
																		  new Microsoft.Office.Common.MethodObject(registerMethod,
																												   registerMethodInvokeType,
																												   false),
																		  new Microsoft.Office.Common.MethodObject(unregisterMethod,
																												   unregisterMethodInvokeType,
																												   false)
																												   );
		this._methodObjectList[eventName]=methodObject;
	},
	unregisterEvent: function (eventName) {
		var e=Function._validateParams(arguments, [
			{ name: "eventName", type: String, mayBeNull: false }
		]);
		if (e) throw e;
		this.unregisterMethod(eventName);
	},
	registerConversation: function Microsoft_Office_Common_ServiceEndPoint$registerConversation(conversationId) {
		var e=Function._validateParams(arguments, [
			{ name: "conversationId", type: String, mayBeNull: false }
			]);
		if (e) throw e;
		this._conversations[conversationId]=true;
	},
	unregisterConversation: function Microsoft_Office_Common_ServiceEndPoint$unregisterConversation(conversationId) {
		var e=Function._validateParams(arguments, [
			{ name: "conversationId", type: String, mayBeNull: false }
			]);
		if (e) throw e;
		delete this._conversations[conversationId];
	},
	setPolicyManager: function Microsoft_Office_Common_ServiceEndPoint$setPolicyManager(policyManager) {
		var e=Function._validateParams(arguments, [
			{ name: "policyManager", type: Object, mayBeNull: false }
			]);
		if (e) throw e;
		if (!policyManager.checkCapability) {
			throw Error.argument("policyManager");
		}
		this._policyManager=policyManager;
	},
	getPolicyManager: function Microsoft_Office_Common_ServiceEndPoint$getPolicyManager() {
		return this._policyManager;
	}
};
Microsoft.Office.Common.ClientEndPoint=function Microsoft_Office_Common_ClientEndPoint(conversationId, targetWindow, targetUrl) {
	var e=Function._validateParams(arguments, [
		  { name: "conversationId", type: String, mayBeNull: false },
		  { name: "targetWindow", mayBeNull: false },
		  { name: "targetUrl", type: String, mayBeNull: false }
	]);
	if (e) throw e;
	if (!targetWindow.postMessage) {
		throw Error.argument("targetWindow");
	}
	this._conversationId=conversationId;
	this._targetWindow=targetWindow;
	this._targetUrl=targetUrl;
	this._callingIndex=0;
	this._callbackList={};
	this._eventHandlerList={};
};
Microsoft.Office.Common.ClientEndPoint.prototype={
	invoke: function Microsoft_Office_Common_ClientEndPoint$invoke(targetMethodName, callback, param) {
		var e=Function._validateParams(arguments, [
			{ name: "targetMethodName", type: String, mayBeNull: false },
			{ name: "callback", type: Function, mayBeNull: true },
			{ name: "param", mayBeNull: true }
		]);
		if (e) throw e;
		var correlationId=this._callingIndex++;
		var now=new Date();
		this._callbackList[correlationId]={"callback" : callback, "createdOn": now.getTime() };
		try {
			var callRequest=new Microsoft.Office.Common.Request(targetMethodName,
																  Microsoft.Office.Common.ActionType.invoke,
																  this._conversationId,
																  correlationId,
																  param);
			var msg=Microsoft.Office.Common.MessagePackager.envelope(callRequest);
			this._targetWindow.postMessage(msg, this._targetUrl);
			Microsoft.Office.Common.XdmCommunicationManager._startMethodTimeoutTimer();
		}
		catch (ex) {
			try {
				if (callback !==null)
					callback(Microsoft.Office.Common.InvokeResultCode.errorInRequest, ex);
			}
			finally {
				delete this._callbackList[correlationId];
			}
		}
	},
	registerForEvent: function Microsoft_Office_Common_ClientEndPoint$registerForEvent(targetEventName, eventHandler, callback, data) {
		var e=Function._validateParams(arguments, [
			{ name: "targetEventName", type: String, mayBeNull: false },
			{ name: "eventHandler", type: Function, mayBeNull: false },
			{ name: "callback", type: Function, mayBeNull: true },
			{ name: "data", mayBeNull: true, optional: true }
		]);
		if (e) throw e;
		var correlationId=this._callingIndex++;
		var now=new Date();
		this._callbackList[correlationId]={"callback" : callback, "createdOn": now.getTime() };
		try {
			var callRequest=new Microsoft.Office.Common.Request(targetEventName,
																  Microsoft.Office.Common.ActionType.registerEvent,
																  this._conversationId,
																  correlationId,
																  data);
			var msg=Microsoft.Office.Common.MessagePackager.envelope(callRequest);
			this._targetWindow.postMessage(msg, this._targetUrl);
			Microsoft.Office.Common.XdmCommunicationManager._startMethodTimeoutTimer();
			this._eventHandlerList[targetEventName]=eventHandler;
		}
		catch (ex) {
			try {
				if (callback !==null) {
					callback(Microsoft.Office.Common.InvokeResultCode.errorInRequest, ex);
				}
			}
			finally {
				delete this._callbackList[correlationId];
			}
		}
	},
	unregisterForEvent: function Microsoft_Office_Common_ClientEndPoint$unregisterForEvent(targetEventName, callback, data) {
		var e=Function._validateParams(arguments, [
			{ name: "targetEventName", type: String, mayBeNull: false },
			{ name: "callback", type: Function, mayBeNull: true },
			{ name: "data", mayBeNull: true, optional: true }
		]);
		if (e) throw e;
		var correlationId=this._callingIndex++;
		var now=new Date();
		this._callbackList[correlationId]={"callback" : callback, "createdOn": now.getTime() };
		try {
			var callRequest=new Microsoft.Office.Common.Request(targetEventName,
																  Microsoft.Office.Common.ActionType.unregisterEvent,
																  this._conversationId,
																  correlationId,
																  data);
			var msg=Microsoft.Office.Common.MessagePackager.envelope(callRequest);
			this._targetWindow.postMessage(msg, this._targetUrl);
			Microsoft.Office.Common.XdmCommunicationManager._startMethodTimeoutTimer();
		}
		catch (ex) {
			try {
				if (callback !==null) {
					callback(Microsoft.Office.Common.InvokeResultCode.errorInRequest, ex);
				}
			}
			finally {
				delete this._callbackList[correlationId];
			}
		}
		finally {
			delete this._eventHandlerList[targetEventName];
		}
	}
};
Microsoft.Office.Common.XdmCommunicationManager=(function () {
	var _invokerQueue=[];
	var _messageProcessingTimer=null;
	var _processInterval=10;
	var _blockingFlag=false;
	var _methodTimeoutTimer=null;
	var _methodTimeout=60000;
	var _serviceEndPoints={};
	var _clientEndPoints={};
	var _initialized=false;
	function _lookupServiceEndPoint(conversationId) {
		for(var id in _serviceEndPoints) {
			 if(_serviceEndPoints[id]._conversations[conversationId]) {
				 return _serviceEndPoints[id];
			 }
		}
		Sys.Debug.trace("Unknown conversation Id.");
		throw Error.argument("conversationId");
	};
	function _lookupClientEndPoint(conversationId) {
		var clientEndPoint=_clientEndPoints[conversationId];
		if(!clientEndPoint) {
			Sys.Debug.trace("Unknown conversation Id.");
			throw Error.argument("conversationId");
		}
		return clientEndPoint;
	};
	function _lookupMethodObject(serviceEndPoint, messageObject) {
		var methodOrEventMethodObject=serviceEndPoint._methodObjectList[messageObject._actionName];
		if (!methodOrEventMethodObject) {
			Sys.Debug.trace("The specified method is not registered on service endpoint:"+messageObject._actionName);
			throw Error.argument("messageObject");
		}
		var methodObject=null;
		if (messageObject._actionType===Microsoft.Office.Common.ActionType.invoke) {
			methodObject=methodOrEventMethodObject;
		} else if (messageObject._actionType===Microsoft.Office.Common.ActionType.registerEvent) {
			methodObject=methodOrEventMethodObject.getRegisterMethodObject();
		} else {
			methodObject=methodOrEventMethodObject.getUnregisterMethodObject();
		}
		return methodObject;
	};
	function _enqueInvoker (invoker) {
		_invokerQueue.push(invoker);
	};
	function _dequeInvoker() {
		if (_messageProcessingTimer !==null) {
			if (!_blockingFlag) {
				if (_invokerQueue.length > 0) {
					var invoker=_invokerQueue.shift();
					_blockingFlag=invoker.getInvokeBlockingFlag();
					invoker.invoke();
				} else {
					clearInterval(_messageProcessingTimer);
					_messageProcessingTimer=null;
				}
			}
		} else {
			Sys.Debug.trace("channel is not ready.");
		}
	};
	function _checkMethodTimeout() {
		if (_methodTimeoutTimer) {
			var clientEndPoint;
			var methodCallsNotTimedout=0;
			var now=new Date();
			for(var conversationId in _clientEndPoints) {
				clientEndPoint=_clientEndPoints[conversationId];
				for(var correlationId in clientEndPoint._callbackList) {
					var callbackEntry=clientEndPoint._callbackList[correlationId];
					if(Math.abs(now.getTime() - callbackEntry.createdOn) >=_methodTimeout) {
						try{
							if(callbackEntry.callback) {
								callbackEntry.callback(Microsoft.Office.Common.InvokeResultCode.errorHandlingMethodCallTimedout, null);
							}
						}
						finally {
							delete clientEndPoint._callbackList[correlationId];
						}
					} else {
						methodCallsNotTimedout++;
					};
				}
			}
			if (methodCallsNotTimedout===0) {
				clearInterval(_methodTimeoutTimer);
				_methodTimeoutTimer=null;
			}
		} else {
			Sys.Debug.trace("channel is not ready.");
		}
	};
	function _postCallbackHandler() {
		_blockingFlag=false;
	};
	function _registerListener(listener) {
		if ((Sys.Browser.agent===Sys.Browser.InternetExplorer) && window.attachEvent) {
			window.attachEvent("onmessage", listener);
		} else if (window.addEventListener) {
			window.addEventListener("message", listener, false);
		} else {
			Sys.Debug.trace("Browser doesn't support the required API.");
			throw Error.argument("Browser");
		}
	};
	function _receive(e) {
		if (e.data !='') {
			var messageObject;
			try {
				messageObject=Microsoft.Office.Common.MessagePackager.unenvelope(e.data);
			}catch(ex) {
				return;
			}
			if ( typeof (messageObject._messageType)=='undefined' ) {
				return;
			}
			if (messageObject._messageType===Microsoft.Office.Common.MessageType.request) {
				var requesterUrl=(e.origin==null || e.origin=="null") ? messageObject._origin : e.origin;
				try {
					var serviceEndPoint=_lookupServiceEndPoint(messageObject._conversationId);
					var policyManager=serviceEndPoint.getPolicyManager();
					if(policyManager && !policyManager.checkCapability(messageObject._conversationId, messageObject._actionName, messageObject._data)) {
						throw "Access Denied";
					}
					var methodObject=_lookupMethodObject(serviceEndPoint, messageObject);
					var invokeCompleteCallback=new Microsoft.Office.Common.InvokeCompleteCallback(e.source,
																										requesterUrl,
																										messageObject._actionName,
																										messageObject._conversationId,
																										messageObject._correlationId,
																										_postCallbackHandler);
					var invoker=new Microsoft.Office.Common.Invoker(methodObject,
																			messageObject._data,
																			invokeCompleteCallback,
																			serviceEndPoint._eventHandlerProxyList,
																			messageObject._conversationId,
																			messageObject._actionName);
					if (_messageProcessingTimer==null) {
						_messageProcessingTimer=setInterval(_dequeInvoker, _processInterval);
					}
					_enqueInvoker(invoker);
				}
				catch (ex) {
					var errorCode=Microsoft.Office.Common.InvokeResultCode.errorHandlingRequest;
					if (ex=="Access Denied") {
						errorCode=Microsoft.Office.Common.InvokeResultCode.errorHandlingRequestAccessDenied;
					}
					var callResponse=new Microsoft.Office.Common.Response(messageObject._actionName,
																				messageObject._conversationId,
																				messageObject._correlationId,
																				errorCode,
																				Microsoft.Office.Common.ResponseType.forCalling,
																				ex);
					var envelopedResult=Microsoft.Office.Common.MessagePackager.envelope(callResponse);
					e.source.postMessage(envelopedResult, requesterUrl);
				}
			} else if (messageObject._messageType===Microsoft.Office.Common.MessageType.response){
				var clientEndPoint=_lookupClientEndPoint(messageObject._conversationId);
				if (messageObject._responseType===Microsoft.Office.Common.ResponseType.forCalling) {
					var callbackEntry=clientEndPoint._callbackList[messageObject._correlationId];
					if (callbackEntry) {
						try {
							if (callbackEntry.callback)
								callbackEntry.callback(messageObject._errorCode, messageObject._data);
						}
						finally {
							delete clientEndPoint._callbackList[messageObject._correlationId];
						}
					}
				} else {
					var eventhandler=clientEndPoint._eventHandlerList[messageObject._actionName];
					if (eventhandler !==undefined && eventhandler !==null) {
						eventhandler(messageObject._data);
					}
				}
			} else {
				return;
			}
		}
	};
	function _initialize () {
		if(!_initialized) {
			_registerListener(_receive);
			_initialized=true;
		}
	};
	return {
		connect : function Microsoft_Office_Common_XdmCommunicationManager$connect(conversationId, targetWindow, targetUrl) {
			_initialize();
			var clientEndPoint=new Microsoft.Office.Common.ClientEndPoint(conversationId, targetWindow, targetUrl);
			_clientEndPoints[conversationId]=clientEndPoint;
			return clientEndPoint;
		},
		getClientEndPoint : function Microsoft_Office_Common_XdmCommunicationManager$getClientEndPoint(conversationId) {
			var e=Function._validateParams(arguments, [
				{name: "conversationId", type: String, mayBeNull: false}
			]);
			if (e) throw e;
			return _clientEndPoints[conversationId];
		},
		createServiceEndPoint : function Microsoft_Office_Common_XdmCommunicationManager$createServiceEndPoint(serviceEndPointId) {
			_initialize();
			var serviceEndPoint=new Microsoft.Office.Common.ServiceEndPoint(serviceEndPointId);
			_serviceEndPoints[serviceEndPointId]=serviceEndPoint;
			return serviceEndPoint;
		},
		getServiceEndPoint : function Microsoft_Office_Common_XdmCommunicationManager$getServiceEndPoint(serviceEndPointId) {
			var e=Function._validateParams(arguments, [
				 {name: "serviceEndPointId", type: String, mayBeNull: false}
			]);
			if (e) throw e;
			return _serviceEndPoints[serviceEndPointId];
		},
		deleteClientEndPoint : function Microsoft_Office_Common_XdmCommunicationManager$deleteClientEndPoint(conversationId) {
			var e=Function._validateParams(arguments, [
				{name: "conversationId", type: String, mayBeNull: false}
			]);
			if (e) throw e;
			delete _clientEndPoints[conversationId];
		},
		_setMethodTimeout : function Microsoft_Office_Common_XdmCommunicationManager$_setMethodTimeout(methodTimeout) {
			var e=Function._validateParams(arguments, [
				{name: "methodTimeout", type: Number, mayBeNull: false}
			]);
			if (e) throw e;
			_methodTimeout=(methodTimeout <=0) ?  60000 : methodTimeout;
		},
		_startMethodTimeoutTimer : function Microsoft_Office_Common_XdmCommunicationManager$_startMethodTimeoutTimer() {
			if (!_methodTimeoutTimer) {
				_methodTimeoutTimer=setInterval(_checkMethodTimeout, _methodTimeout);
			}
		}
	};
})();
Microsoft.Office.Common.Message=function Microsoft_Office_Common_Message(messageType, actionName, conversationId, correlationId, data) {
	var e=Function._validateParams(arguments, [
		{name: "messageType", type: Number, mayBeNull: false},
		{name: "actionName", type: String, mayBeNull: false},
		{name: "conversationId", type: String, mayBeNull: false},
		{name: "correlationId", mayBeNull: false},
		{name: "data", mayBeNull: true, optional: true }
	]);
	if (e) throw e;
	this._messageType=messageType;
	this._actionName=actionName;
	this._conversationId=conversationId;
	this._correlationId=correlationId;
	this._origin=window.location.href;
	if (typeof data=="undefined") {
		this._data=null;
	} else {
		this._data=data;
	}
};
Microsoft.Office.Common.Message.prototype={
	getActionName: function Microsoft_Office_Common_Message$getActionName() {
		return this._actionName;
	},
	getConversationId: function Microsoft_Office_Common_Message$getConversationId() {
		return this._conversationId;
	},
	getCorrelationId: function Microsoft_Office_Common_Message$getCorrelationId() {
		return this._correlationId;
	},
	getOrigin: function Microsoft_Office_Common_Message$getOrigin() {
		return this._origin;
	},
	getData: function Microsoft_Office_Common_Message$getData() {
		return this._data;
	},
	getMessageType: function Microsoft_Office_Common_Message$getMessageType() {
		return this._messageType;
	}
};
Microsoft.Office.Common.Request=function Microsoft_Office_Common_Request(actionName, actionType, conversationId, correlationId, data) {
	Microsoft.Office.Common.Request.uber.constructor.call(this,
														  Microsoft.Office.Common.MessageType.request,
														  actionName,
														  conversationId,
														  correlationId,
														  data);
	this._actionType=actionType;
};
OSF.OUtil.extend(Microsoft.Office.Common.Request, Microsoft.Office.Common.Message);
Microsoft.Office.Common.Request.prototype.getActionType=function Microsoft_Office_Common_Request$getActionType() {
	return this._actionType;
};
Microsoft.Office.Common.Response=function Microsoft_Office_Common_Response(actionName, conversationId, correlationId, errorCode, responseType, data) {
	Microsoft.Office.Common.Response.uber.constructor.call(this,
														   Microsoft.Office.Common.MessageType.response,
														   actionName,
														   conversationId,
														   correlationId,
														   data);
	this._errorCode=errorCode;
	this._responseType=responseType;
};
OSF.OUtil.extend(Microsoft.Office.Common.Response, Microsoft.Office.Common.Message);
Microsoft.Office.Common.Response.prototype.getErrorCode=function Microsoft_Office_Common_Response$getErrorCode() {
	return this._errorCode;
};
Microsoft.Office.Common.Response.prototype.getResponseType=function Microsoft_Office_Common_Response$getResponseType() {
	return this._responseType;
};
Microsoft.Office.Common.MessagePackager={
	envelope: function Microsoft_Office_Common_MessagePackager$envelope(messageObject) {
		return Sys.Serialization.JavaScriptSerializer.serialize(messageObject);
	},
	unenvelope: function Microsoft_Office_Common_MessagePackager$unenvelope(messageObject) {
		return Sys.Serialization.JavaScriptSerializer.deserialize(messageObject);
	}
};
Microsoft.Office.Common.ResponseSender=function Microsoft_Office_Common_ResponseSender(requesterWindow, requesterUrl, actionName, conversationId, correlationId, responseType) {
	var e=Function._validateParams(arguments, [
		{name: "requesterWindow", mayBeNull: false},
		{name: "requesterUrl", type: String, mayBeNull: false},
		{name: "actionName", type: String, mayBeNull: false},
		{name: "conversationId", type: String, mayBeNull: false},
		{name: "correlationId", mayBeNull: false},
		{name: "responsetype", type: Number, maybeNull: false }
		]);
	if (e) throw e;
	this._requesterWindow=requesterWindow;
	this._requesterUrl=requesterUrl;
	this._actionName=actionName;
	this._conversationId=conversationId;
	this._correlationId=correlationId;
	this._invokeResultCode=Microsoft.Office.Common.InvokeResultCode.noError;
	this._responseType=responseType;
	var me=this;
	this._send=function (result) {
		 var response=new Microsoft.Office.Common.Response( me._actionName,
															  me._conversationId,
															  me._correlationId,
															  me._invokeResultCode,
															  me._responseType,
															  result);
		var envelopedResult=Microsoft.Office.Common.MessagePackager.envelope(response);
		me._requesterWindow.postMessage(envelopedResult, me._requesterUrl);
	};
};
Microsoft.Office.Common.ResponseSender.prototype={
	getRequesterWindow: function Microsoft_Office_Common_ResponseSender$getRequesterWindow() {
		return this._requesterWindow;
	},
	getRequesterUrl: function Microsoft_Office_Common_ResponseSender$getRequesterUrl() {
		return this._requesterUrl;
	},
	getActionName: function Microsoft_Office_Common_ResponseSender$getActionName() {
		return this._actionName;
	},
	getConversationId: function Microsoft_Office_Common_ResponseSender$getConversationId() {
		return this._conversationId;
	},
	getCorrelationId: function Microsoft_Office_Common_ResponseSender$getCorrelationId() {
		return this._correlationId;
	},
	getSend: function Microsoft_Office_Common_ResponseSender$getSend() {
		return this._send;
	},
	setResultCode: function Microsoft_Office_Common_ResponseSender$setResultCode(resultCode) {
		this._invokeResultCode=resultCode;
	}
};
Microsoft.Office.Common.InvokeCompleteCallback=function Microsoft_Office_Common_InvokeCompleteCallback(requesterWindow, requesterUrl, actionName, conversationId, correlationId, postCallbackHandler) {
	Microsoft.Office.Common.InvokeCompleteCallback.uber.constructor.call(this,
																 requesterWindow,
																 requesterUrl,
																 actionName,
																 conversationId,
																 correlationId,
																 Microsoft.Office.Common.ResponseType.forCalling);
	this._postCallbackHandler=postCallbackHandler;
	var me=this;
	this._send=function (result) {
		var response=new Microsoft.Office.Common.Response(me._actionName,
															  me._conversationId,
															  me._correlationId,
															  me._invokeResultCode,
															  me._responseType,
															  result);
		var envelopedResult=Microsoft.Office.Common.MessagePackager.envelope(response);
		me._requesterWindow.postMessage(envelopedResult, me._requesterUrl);
		 me._postCallbackHandler();
	};
};
OSF.OUtil.extend(Microsoft.Office.Common.InvokeCompleteCallback, Microsoft.Office.Common.ResponseSender);
Microsoft.Office.Common.Invoker=function Microsoft_Office_Common_Invoker(methodObject, paramValue, invokeCompleteCallback, eventHandlerProxyList, conversationId, eventName) {
	var e=Function._validateParams(arguments, [
		{name: "methodObject", mayBeNull: false},
		{name: "paramValue", mayBeNull: true},
		{name: "invokeCompleteCallback", mayBeNull: false},
		{name: "eventHandlerProxyList", mayBeNull: true},
		{name: "conversationId", type: String, mayBeNull: false},
		{name: "eventName", type: String, mayBeNull: false}
	]);
	if (e) throw e;
	this._methodObject=methodObject;
	this._param=paramValue;
	this._invokeCompleteCallback=invokeCompleteCallback;
	this._eventHandlerProxyList=eventHandlerProxyList;
	this._conversationId=conversationId;
	this._eventName=eventName;
};
Microsoft.Office.Common.Invoker.prototype={
	invoke: function Microsoft_Office_Common_Invoker$invoke() {
		try {
			var result;
			switch (this._methodObject.getInvokeType()) {
				case Microsoft.Office.Common.InvokeType.async:
					this._methodObject.getMethod()(this._param, this._invokeCompleteCallback.getSend());
					break;
				case Microsoft.Office.Common.InvokeType.sync:
					result=this._methodObject.getMethod()(this._param);
					this._invokeCompleteCallback.getSend()(result);
					break;
				case Microsoft.Office.Common.InvokeType.syncRegisterEvent:
					var eventHandlerProxy=this._createEventHandlerProxyObject(this._invokeCompleteCallback);
					result=this._methodObject.getMethod()(eventHandlerProxy.getSend(), this._param);
					this._eventHandlerProxyList[this._conversationId+this._eventName]=eventHandlerProxy.getSend();
					this._invokeCompleteCallback.getSend()(result);
					break;
				case Microsoft.Office.Common.InvokeType.syncUnregisterEvent:
					var eventHandler=this._eventHandlerProxyList[this._conversationId+this._eventName];
					result=this._methodObject.getMethod()(eventHandler, this._param);
					delete this._eventHandlerProxyList[this._conversationId+this._eventName];
					this._invokeCompleteCallback.getSend()(result);
					break;
				case Microsoft.Office.Common.InvokeType.asyncRegisterEvent:
					var eventHandlerProxyAsync=this._createEventHandlerProxyObject(this._invokeCompleteCallback);
					this._methodObject.getMethod()(eventHandlerProxyAsync.getSend(),
												   this._invokeCompleteCallback.getSend(),
												   this._param
												   );
					this._eventHandlerProxyList[this._callerId+this._eventName]=eventHandlerProxyAsync.getSend();
					break;
				case Microsoft.Office.Common.InvokeType.asyncUnregisterEvent:
					var eventHandlerAsync=this._eventHandlerProxyList[this._callerId+this._eventName];
					this._methodObject.getMethod()(eventHandlerAsync,
												   this._invokeCompleteCallback.getSend(),
												   this._param
												   );
					delete this._eventHandlerProxyList[this._callerId+this._eventName];
					break;
				default:
					break;
			}
		}
		catch (ex) {
			this._invokeCompleteCallback.setResultCode(Microsoft.Office.Common.InvokeResultCode.errorInResponse);
			this._invokeCompleteCallback.getSend()(ex);
		}
	},
	getInvokeBlockingFlag: function Microsoft_Office_Common_Invoker$getInvokeBlockingFlag() {
		return this._methodObject.getBlockingFlag();
	},
	_createEventHandlerProxyObject: function Microsoft_Office_Common_Invoker$_createEventHandlerProxyObject(invokeCompleteObject) {
		return new Microsoft.Office.Common.ResponseSender(invokeCompleteObject.getRequesterWindow(),
														  invokeCompleteObject.getRequesterUrl(),
														  invokeCompleteObject.getActionName(),
														  invokeCompleteObject.getConversationId(),
														  invokeCompleteObject.getCorrelationId(),
														  Microsoft.Office.Common.ResponseType.forEventing
														  );
	}
};
OSF.OUtil.setNamespace("OSF", window);
OSF.ClientMode={
	ReadOnly: 0,
	ReadWrite: 1
};
OSF.HostCallPerfMarker={
	IssueCall: "Agave.HostCall.IssueCall",
	ReceiveResponse: "Agave.HostCall.RecieveResponse"
};
OSF.OfficeAppContext=function OSF_OfficeAppContext(id, appName, appVersion, appUILocale, dataLocale, docUrl, clientMode, settings, reason, osfControlType, eToken) {
	this._id=id;
	this._appName=appName;
	this._appVersion=appVersion;
	this._appUILocale=appUILocale;
	this._dataLocale=dataLocale;
	this._docUrl=docUrl;
	this._clientMode=clientMode;
	this._settings=settings;
	this._reason=reason;
	this._osfControlType=osfControlType;
	this._eToken=eToken;
	this.get_id=function get_id() { return this._id; };
	this.get_appName=function get_appName() { return this._appName; };
	this.get_appVersion=function get_appVersion() { return this._appVersion; };
	this.get_appUILocale=function get_appUILocale() { return this._appUILocale; };
	this.get_dataLocale=function get_dataLocale() { return this._dataLocale; };
	this.get_docUrl=function get_docUrl() { return this._docUrl; };
	this.get_clientMode=function get_clientMode() { return this._clientMode; };
	this.get_bindings=function get_bindings() { return this._bindings; };
	this.get_settings=function get_settings() { return this._settings; };
	this.get_reason=function get_reason() { return this._reason; };
	this.get_osfControlType=function get_osfControlType() { return this._osfControlType; };
	this.get_eToken=function get_eToken() { return this._eToken; };
};
OSF.AppName={
	Unsupported: 0,
	Excel: 1,
	Word: 2,
	PowerPoint: 4,
	Outlook: 8,
	ExcelWebApp: 16,
	WordWebApp: 32,
	OutlookWebApp: 64,
	Project: 128
};
OSF.OsfControlType={
	DocumentLevel: 0,
	ContainerLevel: 1
};
OSF.NamespaceManager=function OSF_NamespaceManager(useShortcut) {
	if (useShortcut) {
		this.enableShortcut();
	}
};
OSF.NamespaceManager.prototype={
	enableShortcut: function OSF_NamespaceManager$enableShortcut() {
		if (!this._useShortcut) {
			if (window.Office) {
				this._userOffice=window.Office;
			} else {
				OSF.OUtil.setNamespace("Office", window);
			}
			window.Office=Microsoft.Office.WebExtension;
			this._useShortcut=true;
		}
	},
	disableShortcut: function OSF_NamespaceManager$disableShortcut() {
		if (this._useShortcut) {
			if (this._userOffice) {
				window.Office=this._userOffice;
			} else {
				OSF.OUtil.unsetNamespace("Office", window);
			}
			this._useShortcut=false;
		}
	}
};
OSF.OUtil.setNamespace("DDA", OSF);
OSF.DDA.DocumentMode={
	ReadOnly: 0,
	ReadWrite: 1
};
OSF.OUtil.setNamespace("AsyncResultEnum", OSF.DDA);
OSF.DDA.AsyncResultEnum.Properties={
	Context: "Context",
	Value: "Value",
	Status: "Status",
	Error: "Error"
};
OSF.DDA.AsyncResultEnum.ErrorProperties={
	Name: "Name",
	Message: "Message"
};
OSF.OUtil.setNamespace("Microsoft", window);
OSF.OUtil.setNamespace("Office", Microsoft);
OSF.OUtil.setNamespace("Client", Microsoft.Office);
OSF.OUtil.setNamespace("WebExtension", Microsoft.Office);
Microsoft.Office.WebExtension.InitializationReason={
	Inserted: "inserted",
	DocumentOpened: "documentOpened"
};
Microsoft.Office.WebExtension.ApplicationMode={
	WebEditor: "webEditor",
	WebViewer: "webViewer",
	Client: "client"
};
Microsoft.Office.WebExtension.CoercionType={
	Text: "text",
	Matrix: "matrix",
	Table: "table",
	Html: "html",
	Ooxml: "ooxml"
};
Microsoft.Office.WebExtension.ValueFormat={
	Unformatted: "unformatted",
	Formatted: "formatted"
};
Microsoft.Office.WebExtension.FilterType={
	All: "all",
	OnlyVisible: "onlyVisible"
};
Microsoft.Office.WebExtension.BindingType={
	Text: "text",
	Matrix: "matrix",
	Table: "table"
};
Microsoft.Office.WebExtension.EventType={
	DocumentSelectionChanged: "documentSelectionChanged",
	BindingSelectionChanged: "bindingSelectionChanged",
	BindingDataChanged: "bindingDataChanged",
	DocumentOpened: "documentOpened",
	DocumentClosed: "documentClosed"
};
Microsoft.Office.WebExtension.AsyncResultStatus={
	Succeeded: "succeeded",
	Failed: "failed"
};
Microsoft.Office.WebExtension.OptionalParameters={
	CoercionType: "coercionType",
	ValueFormat: "valueFormat",
	FilterType: "filterType",
	Id: "id",
	PromptText: "promptText",
	StartRow: "startRow",
	StartColumn: "startColumn",
	RowCount: "rowCount",
	ColumnCount: "columnCount",
	Callback: "callback",
	AsyncContext: "asyncContext"
};
OSF.DDA.AsyncResultEnum.ErrorCode={
	Success: 0,
	Failed: 1
};
OSF.DDA.getXdmEventName=function OSF_DDA$GetXdmEventName(bindingId, eventType) {
	if (eventType==Microsoft.Office.WebExtension.EventType.BindingSelectionChanged || eventType==Microsoft.Office.WebExtension.EventType.BindingDataChanged) {
		return bindingId+"_"+eventType;
	} else {
		return eventType;
	}
}
var count=64;
OSF.DDA.MethodDispId={
	dispidMethodMin: count,
	dispidGetSelectedDataMethod: count++,
	dispidSetSelectedDataMethod: count++,
	dispidAddBindingFromSelectionMethod: count++,
	dispidAddBindingFromPromptMethod: count++,
	dispidGetBindingMethod: count++,
	dispidReleaseBindingMethod: count++,
	dispidGetBindingDataMethod: count++,
	dispidSetBindingDataMethod: count++,
	dispidAddRowsMethod: count++,
	dispidClearAllRowsMethod: count++,
	dispidGetAllBindingsMethod: count++,
	dispidLoadSettingsMethod: count++,
	dispidSaveSettingsMethod: count++,
	dispidAddDataPartMethod: count=128,
	dispidGetDataPartByIdMethod:++count,
	dispidGetDataPartsByNamespaceMethod:++count,
	dispidGetDataPartXmlMethod:++count,
	dispidGetDataPartNodesMethod:++count,
	dispidDeleteDataPartMethod:++count,
	dispidGetDataNodeValueMethod:++count,
	dispidGetDataNodeXmlMethod:++count,
	dispidGetDataNodesMethod:++count,
	dispidSetDataNodeValueMethod:++count,
	dispidSetDataNodeXmlMethod:++count,
	dispidAddDataNamespaceMethod:++count,
	dispidGetDataUriByPrefixMethod:++count,
	dispidGetDataPrefixByUriMethod:++count,
	dispidMethodMax: count++};
count=0;
OSF.DDA.EventDispId={
	dispidEventMin: count,
	dispidInitializeEvent: count++,
	dispidSettingsChangedEvent: count++,
	dispidDocumentSelectionChangedEvent: count++,
	dispidBindingSelectionChangedEvent: count++,
	dispidBindingDataChangedEvent: count++,
	dispidDocumentOpenEvent: count++,
	dispidDocumentCloseEvent: count++,
	dispidDataNodeAddedEvent: count=60,
	dispidDataNodeReplacedEvent:++count,
	dispidDataNodeDeletedEvent:++count,
	dispidEventMax:++count
};
OSF.DDA.EventTypeToDispId={};
OSF.DDA.EventTypeToDispId[Microsoft.Office.WebExtension.EventType.DocumentSelectionChanged]=OSF.DDA.EventDispId.dispidDocumentSelectionChangedEvent;
OSF.DDA.EventTypeToDispId[Microsoft.Office.WebExtension.EventType.DocumentOpened]=OSF.DDA.EventDispId.dispidDocumentOpenEvent;
OSF.DDA.EventTypeToDispId[Microsoft.Office.WebExtension.EventType.DocumentClosed]=OSF.DDA.EventDispId.dispidDocumentCloseEvent;
OSF.DDA.EventTypeToDispId[Microsoft.Office.WebExtension.EventType.BindingSelectionChanged]=OSF.DDA.EventDispId.dispidBindingSelectionChangedEvent;
OSF.DDA.EventTypeToDispId[Microsoft.Office.WebExtension.EventType.BindingDataChanged]=OSF.DDA.EventDispId.dispidBindingDataChangedEvent;
OSF.EventDispatch=function OSF_EventDispatch(eventTypes) {
	this._eventHandlers={};
	for(var entry in eventTypes) {
		var eventType=eventTypes[entry];
		this._eventHandlers[eventType]=[];
	}
};
OSF.EventDispatch.prototype={
	getSupportedEvents: function OSF_EventDispatch$getSupportedEvents() {
		var events=[];
		for(var eventName in this._eventHandlers)
			events.push(eventName);
		return events;
	},
	hasEventHandler: function OSF_EventDispatch$hasEventHandler(eventType, handler) {
		var handlers=this._eventHandlers[eventType];
		if(handlers && handlers.length > 0) {
			for(var h in handlers) {
				if(handlers[h]===handler)
					return true;
			}
		}
		return false;
	},
	addEventHandler: function OSF_EventDispatch$addEventHandler(eventType, handler) {
		var handlers=this._eventHandlers[eventType];
		if( handlers && !this.hasEventHandler(eventType, handler) ) {
			handlers.push(handler);
			return true;
		} else {
			return false;
		}
	},
	removeEventHandler: function OSF_EventDispatch$removeEventHandler(eventType, handler) {
		var handlers=this._eventHandlers[eventType];
		if(handlers && handlers.length > 0) {
			for(var index=0; index < handlers.length; index++) {
				if(handlers[index]===handler) {
					handlers.splice(index, 1);
					return true;
				}
			}
		}
		return false;
	},
	clearEventHandlers: function OSF_EventDispatch$clearEventHandlers(eventType) {
		this._eventHandlers[eventType]=[];
	},
	getEventHandlerCount: function OSF_EventDispatch$getEventHandlerCount(eventType) {
		return this._eventHandlers[eventType] !=undefined ? this._eventHandlers[eventType].length : -1;
	},
	fireEvent: function OSF_EventDispatch$fireEvent(eventArgs) {
		if( eventArgs.type==undefined )
			return false;
		var eventType=eventArgs.type;
		if( eventType && this._eventHandlers[eventType] ) {
			var eventHandlers=this._eventHandlers[eventType];
			for(var handler in eventHandlers)
				eventHandlers[handler](eventArgs);
			return true;
		} else {
			return false;
		}
	}
};
OSF.DDA.DataCoercion=(function OSF_DDA_DataCoercion() {
	return {
		getCoercionDefaultForBinding: function OSF_DDA_DataCoercion$getCoercionDefaultForBinding(bindingType) {
			switch(bindingType) {
				case Microsoft.Office.WebExtension.BindingType.Matrix: return Microsoft.Office.WebExtension.CoercionType.Matrix;
				case Microsoft.Office.WebExtension.BindingType.Table: return Microsoft.Office.WebExtension.CoercionType.Table;
				case Microsoft.Office.WebExtension.BindingType.Text:
				default:
					return Microsoft.Office.WebExtension.CoercionType.Text;
			}
		},
		getBindingDefaultForCoercion: function OSF_DDA_DataCoercion$getBindingDefaultForCoercion(coercionType) {
			switch(coercionType) {
				case Microsoft.Office.WebExtension.CoercionType.Matrix: return Microsoft.Office.WebExtension.BindingType.Matrix;
				case Microsoft.Office.WebExtension.CoercionType.Table: return Microsoft.Office.WebExtension.BindingType.Table;
				case Microsoft.Office.WebExtension.CoercionType.Text:
				case Microsoft.Office.WebExtension.CoercionType.Html:
				case Microsoft.Office.WebExtension.CoercionType.Ooxml:
				default:
					return Microsoft.Office.WebExtension.BindingType.Text;
			}
		},
		determineCoercionType: function OSF_DDA_DataCoercion$determineCoercionType(data) {
			if(data==null || data==undefined)
				return null;
			var sourceType=null;
			if(data.rows) {
				sourceType=Microsoft.Office.WebExtension.CoercionType.Table;
			} else if(typeof(data) !="string" && (data.length==0 || data[0] !=undefined)) {
				sourceType=Microsoft.Office.WebExtension.CoercionType.Matrix;
			} else if(typeof(data)=="string") {
				sourceType=Microsoft.Office.WebExtension.CoercionType.Text;
			}
			return sourceType;
		},
		coerceData: function OSF_DDA_DataCoercion$coerceData(data, destinationType, sourceType) {
			sourceType=sourceType || OSF.DDA.DataCoercion.determineCoercionType(data);
			if( sourceType==destinationType ) {
				return data;
			} else {
				return OSF.DDA.DataCoercion._coerceDataFromTable(
					destinationType,
					OSF.DDA.DataCoercion._coerceDataToTable(data, sourceType)
				);
			}
		},
		_matrixToText: function OSF_DDA_DataCoercion$_matrixToText(matrix) {
			if (matrix.length==1 && matrix[0].length==1)
				return ""+matrix[0][0];
			var val="";
			for (var i=0; i < matrix.length; i++) {
				val+=matrix[i].join("\t")+"\n";
			}
			return val.substring(0, val.length - 1);
		},
		_textToMatrix: function OSF_DDA_DataCoercion$_textToMatrix(text) {
			var ret=text.split("\n");
			for (var i=0; i < ret.length; i++)
				ret[i]=ret[i].split("\t");
			return ret;
		},
		_tableToText: function OSF_DDA_DataCoercion$_tableToText(table) {
			var headers="";
			if(table.headers !=null) {
				headers=OSF.DDA.DataCoercion._matrixToText([table.headers])+"\n";
			}
			var rows=OSF.DDA.DataCoercion._matrixToText(table.rows);
			if(rows=="") {
				headers=headers.substring(0, headers.length - 1);
			}
			return headers+rows;
		},
		_tableToMatrix: function OSF_DDA_DataCoercion$_tableToMatrix(table) {
			var matrix=table.rows;
			if(table.headers !=null) {
				matrix.unshift(table.headers);
			}
			return matrix;
		},
		_coerceDataFromTable: function OSF_DDA_DataCoercion$_coerceDataFromTable(coercionType, table) {
			var value;
			switch(coercionType) {
				case Microsoft.Office.WebExtension.CoercionType.Table:
					value=table;
					break;
				case Microsoft.Office.WebExtension.CoercionType.Matrix:
					value=OSF.DDA.DataCoercion._tableToMatrix(table);
					break;
				case Microsoft.Office.WebExtension.CoercionType.Text:
				case Microsoft.Office.WebExtension.CoercionType.Html:
				case Microsoft.Office.WebExtension.CoercionType.Ooxml:
				default:
					value=OSF.DDA.DataCoercion._tableToText(table);
					break;
			}
			return value;
		},
		_coerceDataToTable: function OSF_DDA_DataCoercion$_coerceDataToTable(data, sourceType) {
			if( sourceType==undefined ) {
				sourceType=OSF.DDA.DataCoercion.determineCoercionType(data);
			}
			var value;
			switch(sourceType) {
				case Microsoft.Office.WebExtension.CoercionType.Table:
					value=data;
					break;
				case Microsoft.Office.WebExtension.CoercionType.Matrix:
					value=new Microsoft.Office.WebExtension.TableData(data);
					break;
				case Microsoft.Office.WebExtension.CoercionType.Text:
				case Microsoft.Office.WebExtension.CoercionType.Html:
				case Microsoft.Office.WebExtension.CoercionType.Ooxml:
				default:
					value=new Microsoft.Office.WebExtension.TableData(OSF.DDA.DataCoercion._textToMatrix(data));
					break;
			}
			return value;
		}
	};
})();
OSF.DDA.Context=function OSF_DDA_Context(application, docContainer, document, settings, license) {
	Object.defineProperty(this, "application", {
		value: application,
		writeable: false,
		configurable: false
	});
	if(settings) {
		Object.defineProperty(this, "settings", {
			value: settings,
			writeable: false,
			configurable: false
		});
	}
	if (docContainer) {
		Object.defineProperty(this, "documentContainer", {
			value: docContainer,
			writeable: false,
			configurable: false
		});
	}
	if (document) {
		Object.defineProperty(this, "document", {
			value: document,
			writeable: false,
			configurable: false
		});
	}
	if(license) {
		Object.defineProperty(this, "license", {
			value: license,
			writeable: false,
			configurable: false
		});
	}
}
Object.defineProperty(Microsoft.Office.WebExtension, "context", {
	get: function Microsoft_Office_WebExtension$GetContext() {
		var context;
		if (OSF && OSF._OfficeAppFactory) {
			context=OSF._OfficeAppFactory.getContext();
		}
		return context;
	}
});
Microsoft.Office.WebExtension.useShortNamespace=function Microsoft_Office_WebExtension_useShortNamespace(useShortcut) {
	var namespaceManager=OSF._OfficeAppFactory.getNamespaceManager();
	if(useShortcut) {
		namespaceManager.enableShortcut();
	} else {
		namespaceManager.disableShortcut();
	}
};
Microsoft.Office.WebExtension.select=function Microsoft_Office_WebExtension_select(str) {
	var promise;
	if(str) {
		var index=str.indexOf("#");
		if(index !=-1) {
			var op=str.substring(0, index);
			var target=str.substring(index+1);
			switch(op) {
				case "bindings":
					if(target) {
						promise=new OSF.DDA.BindingPromise(target);
					}
					break;
			}
		}
	}
	if(!promise) {
		throw OSF.OUtil.formatString(Strings.OfficeOM.L_BadSelectorString);
	} else {
		return promise;
	}
};
OSF.DDA.BindingPromise=function OSF_DDA_BindingPromise(bindingId, binding) {
	this._id=bindingId;
	this._binding=binding;
};
OSF.DDA.BindingPromise.prototype={
	_fetch: function OSF_DDA_BindingPromise$_fetch(onComplete, onFail) {
		if(this._binding) {
			if(onComplete)
				onComplete(this._binding);
		} else {
			if(!this._binding) {
				var me=this;
				Microsoft.Office.WebExtension.context.document.bindings.getByIdAsync(this._id, function(asyncResult) {
					if(asyncResult.status==Microsoft.Office.WebExtension.AsyncResultStatus.Succeeded) {
						me._binding=asyncResult.value;
						if(onComplete)
							onComplete(me._binding);
					} else {
						if(onFail)
							onFail(asyncResult);
					}
				});
			}
		}
		return this;
	},
	_onFail: function OSF_DDA_BindingPromise$_onFail(args) {
	},
	getDataAsync: function OSF_DDA_BindingPromise$getDataAsync() {
		var args=arguments;
		this._fetch(function onComplete(binding) { binding.getDataAsync.apply(binding, args); });
		return this;
	},
	setDataAsync: function OSF_DDA_BindingPromise$setDataAsync() {
		var args=arguments;
		this._fetch(function onComplete(binding) { binding.setDataAsync.apply(binding, args); });
		return this;
	},
	addHandlerAsync: function OSF_DDA_BindingPromise$addHandlerAsync() {
		var args=arguments;
		this._fetch(function onComplete(binding) { binding.addHandlerAsync.apply(binding, args); });
		return this;
	},
	removeHandlerAsync: function OSF_DDA_BindingPromise$removeHandlerAsync() {
		var args=arguments;
		this._fetch(function onComplete(binding) { binding.removeHandlerAsync.apply(binding, args); });
		return this;
	}
};
OSF.DDA.License=function OSF_DDA_License(eToken) {
	Object.defineProperty(this, "value", {
		value: eToken,
		writeable: false,
		configurable: false
	});
}
OSF.DDA.Settings=function OSF_DDA_Settings() {};
OSF.DDA.Settings.prototype={
	get : function OSF_DDA_Settings$get(name) {
		var e=Function._validateParams(arguments, [
			{ name: "name", type: String, mayBeNull: false }
		]);
		if (e) throw e;
		var setting=this[name];
		return setting||null;
	},
	set : function OSF_DDA_Settings$set(name, value) {
		var e=Function._validateParams(arguments, [
			{ name: "name", type: String, mayBeNull: false },
			{ name: "value", mayBeNull: true }
		]);
		if (e) throw e;
		this[name]=value;
	},
	remove : function OSF_DDA_Settings$remove(name) {
		var e=Function._validateParams(arguments, [
			{ name: "name", type: String, mayBeNull: false }
		]);
		if (e) throw e;
		delete this[name];
	},
	refreshAsync: function OSF_DDA_Settings$refreshAsync(options) {
		throw OSF.OUtil.formatString(Strings.OfficeOM.L_NotImplemented, 'refreshSettingsAsync');
	},
	saveAsync: function OSF_DDA_Settings$saveAsync(options) {
		throw OSF.OUtil.formatString(Strings.OfficeOM.L_NotImplemented, 'saveSettingsAsync');
	},
	_getSerializedSettings : function OSF_DDA_Settings$_getSerializedSettings() {
		var r={};
		for(var p in this) {
			if(this.hasOwnProperty(p)) {
				try {
					if(JSON) {
						r[p]=JSON.stringify(this[p]);
					} else {
						r[p]=Sys.Serialization.JavaScriptSerializer.serialize(this[p]);
					}
				}
				catch (ex) {
				}
			}
		}
		return r;
	},
	_loadSerializedSettings : function OSF_DDA_Settings$_loadSerializedSettings(settings) {
		settings=settings||{};
		for(var p in settings) {
			if(settings.hasOwnProperty(p)) {
				try {
					if(JSON) {
						this[p]=JSON.parse(settings[p]);
					} else {
						this[p]=Sys.Serialization.JavaScriptSerializer.deserialize(settings[p]);
					}
				}
				catch (ex) {
				}
			}
		}
	}
};
OSF.DDA.Application=function OSF_DDA_Application(officeAppContext, wnd) {
	this._officeAppContext=officeAppContext;
	this._mode=Microsoft.Office.WebExtension.ApplicationMode.Client;
	var getNameString=function (appNameNumber) {
		for (var nameString in OSF.AppName) {
			if (OSF.AppName[nameString]==appNameNumber) return nameString;
		}
		throw OSF.OUtil.formatString(Strings.OfficeOM.L_AppNameNotExist, appNameNumber);
	};
	this._wnd=wnd;
	Object.defineProperties(this, {
		"contentLanguage": {
			value: officeAppContext.get_dataLocale(),
			writeable: false,
			configurable: false
		},
		"displayLanguage": {
			value: officeAppContext.get_appUILocale(),
			writeable: false,
			configurable: false
		},
		"mode": {
			value: officeAppContext.get_clientMode(),
			writeable: false,
			configurable: false
		},
		"name": {
			value: getNameString(officeAppContext.get_appName()),
			writeable: false,
			configurable: false
		},
		"version": {
			value: officeAppContext.get_appVersion(),
			writeable: false,
			configurable: false
		}
	});
};
OSF.DDA.Excel=function OSF_DDA_Excel(officeAppContext, wnd) {
	OSF.DDA.Excel.uber.constructor.call(this, officeAppContext, wnd);
	if (officeAppContext.get_appName()===OSF.AppName.ExcelWebApp) {
		if (officeAppContext.get_clientMode()===OSF.DDA.DocumentMode.ReadOnly) {
			this._mode=Microsoft.Office.WebExtension.ApplicationMode.WebViewer;
		} else {
			this._mode=Microsoft.Office.WebExtension.ApplicationMode.WebEditor;
		}
	}
};
OSF.OUtil.extend(OSF.DDA.Excel, OSF.DDA.Application);
OSF.DDA.Word=function OSF_DDA_Word(officeAppContext, wnd) {
	OSF.DDA.Word.uber.constructor.call(this, officeAppContext, wnd);
};
OSF.OUtil.extend(OSF.DDA.Word, OSF.DDA.Application);
OSF.DDA.Outlook=function OSF_DDA_Outlook(officeAppContext, wnd, outlookAppOm, appReady) {
	OSF.DDA.Outlook.uber.constructor.call(this, officeAppContext, wnd);
	if (officeAppContext.get_appName()===OSF.AppName.OutlookWebApp) {
		this._mode=this._mode=Microsoft.Office.WebExtension.ApplicationMode.WebEditor;
	}
	this.get_outlookAppOm=function() {return outlookAppOm;};
};
OSF.OUtil.extend(OSF.DDA.Outlook, OSF.DDA.Application);
OSF.DDA.OutlookAppOm=function OSF_DDA_OutlookAppOm() {};
OSF.DDA.DocumentContainer=function OSF_DDA_DocumentContainer(officeAppContext, application) {
	this._eventDispatch=new OSF.EventDispatch([
		Microsoft.Office.WebExtension.EventType.DocumentSelectionChanged,
		Microsoft.Office.WebExtension.EventType.DocumentOpened,
		Microsoft.Office.WebExtension.EventType.DocumentClosed
	]);
	Object.defineProperties(this, {
		"application": {
			value: application,
			writeable: false,
			configurable: false
		},
		"mode": {
			value: officeAppContext.get_clientMode(),
			writeable: false,
			configurable: false
		},
		"url": {
			value:  officeAppContext.get_docUrl(),
			writeable: false,
			configurable: false
		}
	});
};
OSF.DDA.DocumentContainer.prototype.getActiveSelectedDataAsync=function OSF_DDA_DocumentContainer$getActiveSelectedDataAsync(coercionType, options) {
	throw OSF.OUtil.formatString(Strings.OfficeOM.L_NotImplemented, 'getActiveSelectedDataAsync');
};
OSF.DDA.DocumentContainer.prototype.setActiveSelectedDataAsync=function OSF_DDA_DocumentContainer$setActiveSelectedDataAsync() {
	throw OSF.OUtil.formatString(Strings.OfficeOM.L_NotImplemented, 'setActiveSelectedDataAsync');
};
OSF.DDA.ExcelContainer=function OSF_DDA_ExcelContainer(officeAppContext, application) {
	throw OSF.OUtil.formatString(Strings.OfficeOM.L_NotImplemented, 'ExcelContainer');
};
OSF.DDA.WordContainer=function OSF_DDA_WordContainer(officeAppContext, application) {
	throw OSF.OUtil.formatString(Strings.OfficeOM.L_NotImplemented, 'WordContainer');
};
OSF.DDA.Document=function OSF_DDA_Document(officeAppContext, application, omFacade, bindingFacade) {
	this._eventDispatch=new OSF.EventDispatch([Microsoft.Office.WebExtension.EventType.DocumentSelectionChanged]);
	this._omFacade=omFacade;
	Object.defineProperties(this, {
		"application": {
			value: application,
			writeable: false,
			configurable: false
		},
		"mode": {
			value: officeAppContext.get_clientMode(),
			writeable: false,
			configurable: false
		},
		"bindings": {
			get: function OSF_DDA_Document$GetBindings() { return bindingFacade; }
		}
	});
	if(!this.url) {
		Object.defineProperty(this, "url", {
			value: officeAppContext.get_docUrl(),
			writeable: false,
			configurable: false
		});
	}
};
OSF.DDA.Document.prototype={
	getSelectedDataAsync: function OSF_DDA_Document$getSelectedDataAsync(coercionType, options) {
		var callback=OSF.OUtil.checkParamsAndGetCallback(arguments, [{ name: "coercionType", type: String }]);
		options=options || {};
		var valueFormat=options[Microsoft.Office.WebExtension.OptionalParameters.ValueFormat] ?
			options[Microsoft.Office.WebExtension.OptionalParameters.ValueFormat] :
			Microsoft.Office.WebExtension.ValueFormat.Unformatted;
		var filter=options[Microsoft.Office.WebExtension.OptionalParameters.FilterType] ?
			options[Microsoft.Office.WebExtension.OptionalParameters.FilterType] :
			Microsoft.Office.WebExtension.FilterType.All;
		this._omFacade.getSelectedData(
			coercionType,
			valueFormat,
			filter,
			callback,
			options[Microsoft.Office.WebExtension.OptionalParameters.AsyncContext]
		);
	},
	setSelectedDataAsync: function OSF_DDA_Document$setSelectedDataAsync(data, options) {
		var callback=OSF.OUtil.checkParamsAndGetCallback(arguments, [{ name: "data" }]);
		options=options || {};
		var sourceType=OSF.DDA.DataCoercion.determineCoercionType(data);
		var coercionType=options[Microsoft.Office.WebExtension.OptionalParameters.CoercionType] || sourceType;
		this._omFacade.setSelectedData(
			coercionType,
			data,
			callback,
			options[Microsoft.Office.WebExtension.OptionalParameters.AsyncContext]
		);
	},
	addHandlerAsync: function OSF_DDA_Document$addHandlerAsync(eventType, handler, options) {
		var callback=OSF.OUtil.checkParamsAndGetCallback(arguments, [
			{ name: "eventType", type: String },
			{ name: "handler", type: Function }
		]);
		if (callback==handler)
			callback=null;
		options=options || {};
		var eventDispatch=this._eventDispatch;
		var afterRegistration=function(succeeded) {
			succeeded=succeeded && eventDispatch.addEventHandler(eventType, handler);
			var errorArgs;
			if(!succeeded) {
				errorArgs={};
				errorArgs[OSF.DDA.AsyncResultEnum.ErrorProperties.Name]=OSF.DDA.AsyncResultEnum.ErrorCode.Failed;
				errorArgs[OSF.DDA.AsyncResultEnum.ErrorProperties.Message]=OSF.OUtil.formatString(Strings.OfficeOM.L_EventHandlerAdditionFailed);
			}
			var asyncInitArgs={};
			asyncInitArgs[OSF.DDA.AsyncResultEnum.Properties.Context]=options[Microsoft.Office.WebExtension.OptionalParameters.AsyncContext];
			asyncInitArgs[OSF.DDA.AsyncResultEnum.Properties.Value]=this;
			if(callback) {
				callback(new OSF.DDA.AsyncResult(asyncInitArgs, errorArgs));
			}
		};
		if(this._eventDispatch.getEventHandlerCount(eventType)==0) {
			this._omFacade.registerEvent("" , eventType, eventDispatch, afterRegistration);
		}
		else {
			afterRegistration(true);
		}
	},
	removeHandlerAsync: function OSF_DDA_Document$removeHandlerAsync(eventType, handler, options) {
		var callback=OSF.OUtil.checkParamsAndGetCallback(arguments, [
			{ name: "eventType", type: String },
			{ name: "handler", type: Function }
		]);
		if (callback==handler)
			callback=null;
		options=options || {};
		var afterUnregistration=function(succeeded) {
			var errorArgs;
			if(!succeeded) {
				errorArgs={};
				errorArgs[OSF.DDA.AsyncResultEnum.ErrorProperties.Name]=OSF.DDA.AsyncResultEnum.ErrorCode.Failed;
				errorArgs[OSF.DDA.AsyncResultEnum.ErrorProperties.Message]=OSF.OUtil.formatString(Strings.OfficeOM.L_EventHandlerRemovalFailed);
			}
			var asyncInitArgs={};
			asyncInitArgs[OSF.DDA.AsyncResultEnum.Properties.Context]=options[Microsoft.Office.WebExtension.OptionalParameters.AsyncContext];
			asyncInitArgs[OSF.DDA.AsyncResultEnum.Properties.Value]=this;
			if(callback) {
				callback(new OSF.DDA.AsyncResult(asyncInitArgs, errorArgs));
			}
		};
		var succeeded;
		if(handler) {
			succeeded=this._eventDispatch.removeEventHandler(eventType, handler);
		} else {
			this._eventDispatch.clearEventHandlers(eventType);
			succeeded=true;
		}
		if(succeeded && this._eventDispatch.getEventHandlerCount(eventType)==0) {
			this._omFacade.unregisterEvent("" , eventType, afterUnregistration);
		}
		else {
			afterUnregistration(succeeded);
		}
	}
};
OSF.DDA.ExcelDocument=function OSF_DDA_ExcelDocument(officeAppContext, application) {
	throw OSF.OUtil.formatString(Strings.OfficeOM.L_NotImplemented, 'ExcelDocument');
};
OSF.DDA.WordDocument=function OSF_DDA_WordDocument(officeAppContext, application) {
	throw OSF.OUtil.formatString(Strings.OfficeOM.L_NotImplemented, 'WordDocument');
};
OSF.DDA.BindingFacade=function OSF_DDA_BindingFacade(docInstance) {
	this._eventDispatches=[];
	Object.defineProperty(this, "document", {
		value: docInstance,
		writeable: false,
		configurable: false
	});
};
OSF.DDA.BindingFacade.prototype={
	_generateBindingId: function OSF_DDA_BindingFacade$_generateBindingId() {
		return "UnnamedBinding_"+OSF.OUtil.getUniqueId()+"_"+new Date().getTime();
	},
	addFromSelectionAsync: function OSF_DDA_BindingFacade$addFromSelectionAsync(bindingType, options) {
		var callback=OSF.OUtil.checkParamsAndGetCallback(arguments, [{ name: "bindingType", type: String }]);
		options=options || {};
		var bindingId=options[Microsoft.Office.WebExtension.OptionalParameters.Id] ?
			options[Microsoft.Office.WebExtension.OptionalParameters.Id] :
			this._generateBindingId();
		this.document._omFacade.addBindingFromSelection(
			bindingId,
			bindingType,
			callback,
			options[Microsoft.Office.WebExtension.OptionalParameters.AsyncContext]
		);
	},
	addFromPromptAsync: function OSF_DDA_BindingFacade$addFromPromptAsync(bindingType, options) {
		var callback=OSF.OUtil.checkParamsAndGetCallback(arguments, [{ name: "bindingType", type: String }]);
		options=options || {};
		var bindingId=options[Microsoft.Office.WebExtension.OptionalParameters.Id] ?
			options[Microsoft.Office.WebExtension.OptionalParameters.Id] :
			this._generateBindingId();
		var promptText=options[Microsoft.Office.WebExtension.OptionalParameters.PromptText] ?
			options[Microsoft.Office.WebExtension.OptionalParameters.PromptText] :
			"Please make a selection";
		this.document._omFacade.addBindingFromPrompt(
			bindingId,
			bindingType,
			promptText,
			callback,
			options[Microsoft.Office.WebExtension.OptionalParameters.AsyncContext]
		);
	},
	getAllAsync: function OSF_DDA_BindingFacade$getAllAsync(options) {
		var callback=OSF.OUtil.checkParamsAndGetCallback(arguments, []);
		options=options || {};
		this.document._omFacade.getAllBindings(callback, options[Microsoft.Office.WebExtension.OptionalParameters.AsyncContext]);
	},
	getByIdAsync: function OSF_DDA_BindingFacade$getByIdAsync(bindingId, options) {
		var callback=OSF.OUtil.checkParamsAndGetCallback(arguments, [{ name: "bindingId", type: String }]);
		options=options || {};
		this.document._omFacade.getBinding(bindingId, callback, options[Microsoft.Office.WebExtension.OptionalParameters.AsyncContext]);
	},
	releaseByIdAsync: function OSF_DDA_BindingFacade$releaseByIdAsync(bindingId, options) {
		var callback=OSF.OUtil.checkParamsAndGetCallback(arguments, [{ name: "bindingId", type: String }]);
		options=options || {};
		this.document._omFacade.releaseBinding(bindingId, callback, options[Microsoft.Office.WebExtension.OptionalParameters.AsyncContext]);
	}
};
OSF.DDA.Binding=function OSF_DDA_Binding(id, docInstance) {
	Object.defineProperties(this, {
		"document": {
			value: docInstance,
			writeable: false,
			configurable: false
		},
		"id": {
			value: id,
			writeable: false,
			configurable: false
		},
		"type": {
			get: function OSF_DDA_Binding$GetType() { throw "not Implemented"; },
			configurable: true
		}
	});
};
OSF.DDA.Binding.prototype={
	getDataAsync: function OSF_DDA_Binding$getDataAsync(options) {
		var callback=OSF.OUtil.checkParamsAndGetCallback(arguments, []);
		options=options || {};
		var coercionType=options[Microsoft.Office.WebExtension.OptionalParameters.CoercionType] || OSF.DDA.DataCoercion.getCoercionDefaultForBinding(this.type);
		var valueFormat=options[Microsoft.Office.WebExtension.OptionalParameters.ValueFormat] || Microsoft.Office.WebExtension.ValueFormat.Unformatted;
		var filter=options[Microsoft.Office.WebExtension.OptionalParameters.FilterType] || Microsoft.Office.WebExtension.FilterType.All;
		var subset=null;
		if(this.rowCount !=undefined && this.columnCount !=undefined) {
			var startRow=options[Microsoft.Office.WebExtension.OptionalParameters.StartRow] || 0;
			var startCol=options[Microsoft.Office.WebExtension.OptionalParameters.StartColumn] || 0;
			var rowCount=options[Microsoft.Office.WebExtension.OptionalParameters.RowCount] || 0;
			var colCount=options[Microsoft.Office.WebExtension.OptionalParameters.ColumnCount] || 0;
			if(!(startRow==startCol && startCol==rowCount && rowCount==colCount && colCount==0)) {
				subset=[
					[startRow, startCol],
					[rowCount, colCount]
				];
			}
		}
		this.document._omFacade.getBindingData(
			this.id,
			coercionType,
			valueFormat,
			filter,
			subset,
			callback,
			options[Microsoft.Office.WebExtension.OptionalParameters.AsyncContext]
		);
	},
	setDataAsync: function OSF_DDA_Binding$setDataAsync(data, options) {
		var callback=OSF.OUtil.checkParamsAndGetCallback(arguments, [{ name: "data" }]);
		options=options || {};
		var coercionType=options[Microsoft.Office.WebExtension.OptionalParameters.CoercionType] || OSF.DDA.DataCoercion.determineCoercionType(data);
		var offset=null;
		if(this.rowCount !=undefined && this.columnCount !=undefined) {
			var startRow=options[Microsoft.Office.WebExtension.OptionalParameters.StartRow] || 0;
			var startCol=options[Microsoft.Office.WebExtension.OptionalParameters.StartColumn] || 0;
			if(!(startRow==startCol && startCol==0)) {
				offset=[startRow, startCol];
			}
		}
		this.document._omFacade.setBindingData(
			this.id,
			coercionType,
			data,
			offset,
			callback,
			options[Microsoft.Office.WebExtension.OptionalParameters.AsyncContext]
		);
	},
	addHandlerAsync: function OSF_DDA_Binding$addHandlerAsync(eventType, handler, options) {
		var callback=OSF.OUtil.checkParamsAndGetCallback(arguments, [
			{ name: "eventType", type: String },
			{ name: "handler", type: Function }
		]);
		if (callback==handler)
			callback=null;
		options=options || {};
		var id=this.id;
		var bindingFacade=this.document.bindings;
		if(!bindingFacade._eventDispatches[id])
			bindingFacade._eventDispatches[id]=new OSF.EventDispatch([Microsoft.Office.WebExtension.EventType.BindingSelectionChanged, Microsoft.Office.WebExtension.EventType.BindingDataChanged]);
		var eventDispatch=bindingFacade._eventDispatches[id];
		var afterRegistration=function(succeeded) {
			if(succeeded) {
				succeeded=succeeded && eventDispatch.addEventHandler(eventType, handler);
			}
			var errorArgs;
			if(!succeeded) {
				errorArgs={};
				errorArgs[OSF.DDA.AsyncResultEnum.ErrorProperties.Name]=OSF.DDA.AsyncResultEnum.ErrorCode.Failed;
				errorArgs[OSF.DDA.AsyncResultEnum.ErrorProperties.Message]=OSF.OUtil.formatString(Strings.OfficeOM.L_EventHandlerAdditionFailed);
			}
			var asyncInitArgs={};
			asyncInitArgs[OSF.DDA.AsyncResultEnum.Properties.Context]=options[Microsoft.Office.WebExtension.OptionalParameters.AsyncContext];
			asyncInitArgs[OSF.DDA.AsyncResultEnum.Properties.Value]=this;
			if(callback) {
				callback(new OSF.DDA.AsyncResult(asyncInitArgs, errorArgs));
			}
		}
		if(eventDispatch.getEventHandlerCount(eventType)==0) {
			this.document._omFacade.registerEvent(id, eventType, eventDispatch, afterRegistration);
		}
		else {
			afterRegistration(true);
		}
	},
	removeHandlerAsync: function OSF_DDA_Binding$removeHandlerAsync(eventType, handler, options) {
		var callback=OSF.OUtil.checkParamsAndGetCallback(arguments, [
			{ name: "eventType", type: String },
			{ name: "handler", type: Function }
		]);
		if (callback==handler)
			callback=null;
		options=options || {};
		var id=this.id;
		var afterUnregistration=function(succeeded) {
			var errorArgs;
			if(!succeeded) {
				errorArgs={};
				errorArgs[OSF.DDA.AsyncResultEnum.ErrorProperties.Name]=OSF.DDA.AsyncResultEnum.ErrorCode.Failed;
				errorArgs[OSF.DDA.AsyncResultEnum.ErrorProperties.Message]=OSF.OUtil.formatString(Strings.OfficeOM.L_EventHandlerRemovalFailed);
			}
			var asyncInitArgs={};
			asyncInitArgs[OSF.DDA.AsyncResultEnum.Properties.Context]=options[Microsoft.Office.WebExtension.OptionalParameters.AsyncContext];
			asyncInitArgs[OSF.DDA.AsyncResultEnum.Properties.Value]=this;
			if(callback) {
				callback(new OSF.DDA.AsyncResult(asyncInitArgs, errorArgs));
			}
		};
		var succeeded=false;
		var bindingFacade=this.document.bindings;
		if(bindingFacade._eventDispatches[id]) {
			var eventDispatch=bindingFacade._eventDispatches[id];
			if(handler) {
				succeeded=eventDispatch.removeEventHandler(eventType, handler);
			} else {
				eventDispatch.clearEventHandlers(eventType);
				succeeded=true;
			}
			if(succeeded && eventDispatch.getEventHandlerCount(eventType)==0) {
				this.document._omFacade.unregisterEvent(id, eventType, afterUnregistration);
			}
			else {
				afterUnregistration(succeeded);
			}
		}
	}
};
OSF.DDA.TextBinding=function OSF_DDA_TextBinding(id, docInstance) {
	OSF.DDA.TextBinding.uber.constructor.call(
		this,
		id,
		docInstance
	);
	Object.defineProperty(this, "type", {
		value: Microsoft.Office.WebExtension.BindingType.Text,
		writeable: false,
		configurable: false
	});
};
OSF.OUtil.extend(OSF.DDA.TextBinding, OSF.DDA.Binding);
OSF.DDA.MatrixBinding=function OSF_DDA_MatrixBinding(id, docInstance, rows, cols) {
	OSF.DDA.MatrixBinding.uber.constructor.call(
		this,
		id,
		docInstance
	);
	Object.defineProperties(this, {
		"type": {
			value: Microsoft.Office.WebExtension.BindingType.Matrix,
			writeable: false,
			configurable: false
		},
		"rowCount": {
			value: rows ? rows : 0,
			writeable: false,
			configurable: false
		},
		"columnCount": {
			value: cols ? cols: 0,
			writeable: false,
			configurable: false
		}
	});
};
OSF.OUtil.extend(OSF.DDA.MatrixBinding, OSF.DDA.Binding);
OSF.DDA.TableBinding=function OSF_DDA_TableBinding(id, docInstance, rows, cols, hasHeaders) {
	OSF.DDA.TableBinding.uber.constructor.call(
		this,
		id,
		docInstance
	);
	Object.defineProperties(this, {
		"type": {
			value: Microsoft.Office.WebExtension.BindingType.Table,
			writeable: false,
			configurable: false
		},
		"rowCount": {
			value: rows ? rows : 0,
			writeable: false,
			configurable: false
		},
		"columnCount": {
			value: cols ? cols: 0,
			writeable: false,
			configurable: false
		},
		"hasHeaders": {
			value: hasHeaders ? hasHeaders : false,
			writeable: false,
			configurable: false
		}
	});
};
OSF.OUtil.extend(OSF.DDA.TableBinding, OSF.DDA.Binding);
OSF.DDA.TableBinding.prototype.addRowsAsync=function OSF_DDA_TableBinding$addRowsAsync(rows, options) {
	var callback=OSF.OUtil.checkParamsAndGetCallback(arguments, [{ name: "rows", type: Array }]);
	options=options || {};
	this.document._omFacade.addTableRows(
		this.id,
		rows,
		callback,
		options[Microsoft.Office.WebExtension.OptionalParameters.AsyncContext]
	);
};
OSF.DDA.TableBinding.prototype.deleteAllDataValuesAsync=function OSF_DDA_TableBinding$deleteAllDataValuesAsync(options) {
	var callback=OSF.OUtil.checkParamsAndGetCallback(arguments, []);
	options=options || {};
	this.document._omFacade.clearAllRows(
		this.id,
		callback,
		options[Microsoft.Office.WebExtension.OptionalParameters.AsyncContext]
	);
};
Microsoft.Office.WebExtension.TableData=function Microsoft_Office_WebExtension_TableData(rows, headers) {
	Object.defineProperties(this, {
		"headers": {
			get: function() { return headers ? headers : null; },
			set: function(value) {
				if(typeof value=="object") {
					headers=value;
					return true;
				} else {
					return false;
				}
			}
		},
		"rows": {
			get: function() { return rows ? rows : []; },
			set: function(value) {
				if(typeof value=="object" && (value.length==0 || value[0] !=undefined)) {
					rows=value;
					return true;
				} else {
					return false;
				}
			}
		}
	});
};
OSF.DDA.Error=function OSF_DDA_Error(name, message) {
	Object.defineProperties(this, {
		"name": {
			value: name,
			writeable: false,
			configurable: false
		},
		"message": {
			value: message,
			writeable: false,
			configurable: false
		}
	});
};
OSF.DDA.AsyncResult=function OSF_DDA_AsyncResult(initArgs, errorArgs) {
	Object.defineProperties(this, {
		"value": {
			value: initArgs[OSF.DDA.AsyncResultEnum.Properties.Value],
			writeable: false,
			configurable: false
		},
		"status": {
			value: errorArgs ? Microsoft.Office.WebExtension.AsyncResultStatus.Failed : Microsoft.Office.WebExtension.AsyncResultStatus.Succeeded,
			writeable: false,
			configurable: false
		}
	});
	if(initArgs[OSF.DDA.AsyncResultEnum.Properties.Context]) {
		Object.defineProperty(this, "context", {
			value: initArgs[OSF.DDA.AsyncResultEnum.Properties.Context],
			writeable: false,
			configurable: false
		});
	}
	if(errorArgs) {
		Object.defineProperty(this, "error", {
			value: new OSF.DDA.Error(
				errorArgs[OSF.DDA.AsyncResultEnum.ErrorProperties.Name],
				errorArgs[OSF.DDA.AsyncResultEnum.ErrorProperties.Message]
			),
			writeable: false,
			configurable: false
		});
	}
};
OSF.DDA.DocumentSelectionChangedEventArgs=function OSF_DDA_DocumentSelectionChangedEventArgs(appOm) {
	var appOmName;
	if(appOm.getSelectedDataAsync) {
		appOmName="document";
	}
	if(appOm.getActiveSelectedDataAsync) {
		appOmName="documentContainer";
	}
	Object.defineProperties(this, {
		"type": {
			value: Microsoft.Office.WebExtension.EventType.DocumentSelectionChanged,
			writeable: false,
			configurable: false
		},
		appOmName: {
			value: appOm,
			writeable: false,
			configurable: false
		}
	});
};
OSF.DDA.BindingSelectionChangedEventArgs=function OSF_DDA_BindingSelectionChangedEventArgs(bindingInstance) {
	Object.defineProperties(this, {
		"type": {
			value: Microsoft.Office.WebExtension.EventType.BindingSelectionChanged,
			writeable: false,
			configurable: false
		},
		"binding": {
			value: bindingInstance,
			writeable: false,
			configurable: false
		}
	});
};
OSF.DDA.BindingDataChangedEventArgs=function OSF_DDA_BindingDataChangedEventArgs(bindingInstance) {
	Object.defineProperties(this, {
		"type": {
			value: Microsoft.Office.WebExtension.EventType.BindingDataChanged,
			writeable: false,
			configurable: false
		},
		"binding": {
			value: bindingInstance,
			writeable: false,
			configurable: false
		}
	});
};
OSF.DDA.ActiveSelectionChangedEventArgs=function OSF_DDA_ActiveSelectionChangedEventArgs(docContainer) {
	Object.defineProperties(this, {
		"type": {
			value: Microsoft.Office.WebExtension.EventType.DocumentSelectionChanged,
			writeable: false,
			configurable: false
		},
		"documentContainer": {
			value: docContainer,
			writeable: false,
			configurable: false
		}
	});
};
OSF.DDA.DocumentOpenedEventArgs=function OSF_DDA_DocumentOpenedEventArgs(docContainer) {
	Object.defineProperty(this, "type", {
		value: Microsoft.Office.WebExtension.EventType.DocumentOpened,
		writeable: false,
		configurable: false
	});
};
OSF.DDA.DocumentClosedEventArgs=function OSF_DDA_DocumentClosedEventArgs() {
	Object.defineProperties(this, {
		"type": {
			value: Microsoft.Office.WebExtension.EventType.DocumentClosed,
			writeable: false,
			configurable: false
		}
	});
};
OSF._OfficeAppFactory=(function() {
	var _officeJS="office.js";
	var _appToScriptTable={
		"1-15" : "excel-15.debug.js",
		"2-15" : "word-15.debug.js",
		"8-15" : "outlook-15.debug.js",
		"16-15" : "excelwebapp-15.debug.js",
		"64-15" : "outlookwebapp-15.debug.js",
		"128-15": "Project-15.debug.js"
	};
	var _namespaceManager;
	var _context;
	var _app;
	var _WebAppState={};
	_WebAppState.id=null;
	_WebAppState.webAppUrl=null;
	_WebAppState.conversationID=null;
	_WebAppState.clientEndPoint=null;
	_WebAppState.window=window.parent;
	var _isRichClient=true;
	var retrieveIframeInfo=function () {
		var xdmInfoValue=OSF.OUtil.parseXdmInfo();
		if (xdmInfoValue !=null) {
			var items=xdmInfoValue.split('|');
			if (items !=undefined && items.length==3) {
				_WebAppState.conversationID=items[0];
				_WebAppState.id=items[1];
				_WebAppState.webAppUrl=items[2];
				_isRichClient=false;
			}
		}
	};
	var getAppContextAsync=function (wnd, gotAppContext) {
			if (_isRichClient) {
				var returnedContext;
				var context=window.external.GetContext();
				var appType=context.GetAppType();
				var appTypeSupported=false;
				for (var appEntry in OSF.AppName) {
					if (OSF.AppName[appEntry]==appType) {
						appTypeSupported=true;
						break;
					}
				}
				if (!appTypeSupported) {
					throw "Unsupported client type "+appType;
				}
				var id=context.GetSolutionRef();
				var version=context.GetAppVersionMajor();
				var UILocale=context.GetAppUILocale();
				var dataLocale=context.GetAppDataLocale();
				var docUrl=context.GetDocUrl();
				var clientMode=context.GetAppCapabilities();
				var reason=context.GetActivationMode();
				var osfControlType=context.GetControlIntegrationLevel();
				var eToken;
				try {
					eToken=context.GetSolutionToken();
				}
				catch (ex) {
				}
				eToken=eToken ? eToken.toString() : "";
				var keys=[];
				var values=[];
				context.GetSettings().Read(keys, values);
				var settings={};
				var i=0;
				for (i=0; i< keys.length; i++) {
					settings[keys[i]]=values[i];
				}
				returnedContext=new OSF.OfficeAppContext(id, appType, version, UILocale, dataLocale, docUrl, clientMode, settings, reason, osfControlType, eToken);
				gotAppContext(returnedContext);
			} else {
				var getInvocationCallbackWebApp=function (errorCode, appContext) {
					var settings;
					if (appContext._appName===OSF.AppName.ExcelWebApp) {
						var serializedSettings=appContext._settings;
						settings={};
						for(var index in serializedSettings) {
							var setting=serializedSettings[index];
							settings[setting[0]]=setting[1];
						}
					}
					else {
						settings=appContext._settings;
					}
					if (errorCode===0 && appContext._id !=undefined && appContext._appName !=undefined && appContext._appVersion !=undefined && appContext._appUILocale !=undefined && appContext._dataLocale !=undefined &&
						appContext._docUrl !=undefined && appContext._clientMode !=undefined && appContext._settings !=undefined && appContext._reason !=undefined) {
						var returnedContext=new OSF.OfficeAppContext(appContext._id, appContext._appName, appContext._appVersion, appContext._appUILocale, appContext._dataLocale, appContext._docUrl, appContext._clientMode, settings, appContext._reason, appContext._osfControlType, appContext._eToken);
						gotAppContext(returnedContext);
					} else {
						throw "Function ContextActivationManager_getAppContextAsync call failed. ErrorCode is "+errorCode;
					}
				};
				_WebAppState.clientEndPoint.invoke("ContextActivationManager_getAppContextAsync", getInvocationCallbackWebApp, _WebAppState.id);
			}
		};
	var initialize=function () {
		_namespaceManager=new OSF.NamespaceManager(true );
		retrieveIframeInfo();
		if (!_isRichClient) {
			_WebAppState.clientEndPoint=Microsoft.Office.Common.XdmCommunicationManager.connect(_WebAppState.conversationID, _WebAppState.window, _WebAppState.webAppUrl);
		}
		var scripts=document.getElementsByTagName("script") || [];
		var i, src, basePath, indexOfOfficeJS;
		for (i=0;i<scripts.length;i++) {
			if (scripts[i].src) {
				src=scripts[i].src.toLowerCase();
				indexOfOfficeJS=src.indexOf(_officeJS);
				if (indexOfOfficeJS===(src.length - _officeJS.length) && (indexOfOfficeJS===0 || src.charAt(indexOfOfficeJS-1)==='/' || src.charAt(indexOfOfficeJS-1)==='\\')) {
					basePath=src.replace(_officeJS, "");
				}
			}
		}
		if (basePath===undefined) throw "Office Web Extension script library file name should be Office.js.";
		getAppContextAsync(_WebAppState.window, function (appContext) {
			var app, docContainer, doc;
			var retryNumber=100;
			var t;
			function appReady() {
				if (Microsoft.Office.WebExtension.initialize !=undefined && app !=undefined) {
					var license=new OSF.DDA.License(appContext.get_eToken());
					var settings=new OSF.DDA.Settings();
					settings._loadSerializedSettings(appContext.get_settings());
					if (appContext.get_appName()==OSF.AppName.OutlookWebApp || appContext.get_appName()==OSF.AppName.Outlook) {
						_context=new OSF.DDA.Context(app, null, null, settings, license);
						Microsoft.Office.WebExtension.initialize();
					}
					else {
						if(appContext.get_appName()==OSF.AppName.Project) {
							docContainer=null;
						}
						else if (appContext.get_osfControlType()===OSF.OsfControlType.DocumentLevel) {
							docContainer=null;
						}
						else if (appContext.get_osfControlType()===OSF.OsfControlType.ContainerLevel) {
							docContainer=null;
						}
						else {
							throw OSF.OUtil.formatString(Strings.OfficeOM.L_OsfControlTypeNotSupported);
						}
						_context=new OSF.DDA.Context(app, docContainer, doc, settings, license);
						var reason=appContext.get_reason();
						if(_isRichClient) {
							reason=OSF.DDA.RichInitializationReason[reason];
						}
						Microsoft.Office.WebExtension.initialize(reason);
					}
					_app=app;
					if (t !=undefined) window.clearTimeout(t);
				} else if (retryNumber==0) {
					clearTimeout(t);
					throw OSF.OUtil.formatString(Strings.OfficeOM.L_InitializeNotReady);
				} else {
					retryNumber--;
					t=window.setTimeout(appReady, 100);
				}
			};
			var localeStringFileLoaded=function() {
				if (typeof Strings=='undefined' || typeof Strings.OfficeOM=='undefined') throw "The locale, "+appContext.get_appUILocale()+", provided by the host app is not supported.";
				var scriptPath=basePath+_appToScriptTable[appContext.get_appName()+"-"+appContext.get_appVersion()];
				if (appContext.get_appName()==OSF.AppName.ExcelWebApp || appContext.get_appName()==OSF.AppName.Excel) {
					var excelScriptLoaded=function() {
						app=new OSF.DDA.Excel(appContext, _WebAppState.window);
						doc=new OSF.DDA.ExcelDocument(appContext, app);
						docContainer=new OSF.DDA.ExcelContainer(appContext, app);
						appReady();
					};
					OSF.OUtil.loadScript(scriptPath, excelScriptLoaded);
				} else if (appContext.get_appName()==OSF.AppName.Word) {
					var wordScriptLoaded=function() {
						app=new OSF.DDA.Word(appContext, _WebAppState.window);
						doc=new OSF.DDA.WordDocument(appContext, app);
						docContainer=new OSF.DDA.WordContainer(appContext, app);
						appReady();
					};
					OSF.OUtil.loadScript(scriptPath, wordScriptLoaded);
				} else if (appContext.get_appName()==OSF.AppName.OutlookWebApp || appContext.get_appName()==OSF.AppName.Outlook) {
					var outlookScriptLoaded=function() {
						var outlookAppOm=new OSF.DDA.OutlookAppOm();
						app=new OSF.DDA.Outlook(appContext, _WebAppState.window, outlookAppOm, appReady);
					};
					OSF.OUtil.loadScript(scriptPath, outlookScriptLoaded);
				} else if (appContext.get_appName()==OSF.AppName.Project) {
					var projScriptLoaded=function() {
						app=new OSF.DDA.Project(appContext, _WebAppState.window);
						doc=new OSF.DDA.ProjectDocument(appContext, app);
						appReady();
					};
					OSF.OUtil.loadScript(scriptPath, projScriptLoaded);
				} else {
					throw OSF.OUtil.formatString(Strings.OfficeOM.L_AppNotExistInitializeNotCalled, appContext.get_appName());
				}
			};
			OSF.OUtil.loadScript(basePath+appContext.get_appUILocale()+"/office_strings.debug.js", localeStringFileLoaded);
		});
	};
	initialize();
	return {
		getId : function OSF__OfficeAppFactory$getId() {return _WebAppState.id;},
		getClientEndPoint : function OSF__OfficeAppFactory$getClientEndPoint() { return _WebAppState.clientEndPoint; },
		getApp : function OSF__OfficeAppFactory$getApp() { return _app; },
		getWebAppState : function OSF__OfficeAppFactory$getWebAppState() { return _WebAppState; },
		getNamespaceManager : function OSF__OfficeAppFactory$getNamespaceManager() { return _namespaceManager; },
		getContext: function OSF__OfficeAppFactory$getContext() { return _context; }
	};
})();

