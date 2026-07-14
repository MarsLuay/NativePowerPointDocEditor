import { Fragment as e, Teleport as t, computed as n, createBlock as r, createCommentVNode as i, createElementBlock as a, createElementVNode as o, createTextVNode as s, defineComponent as c, nextTick as l, normalizeClass as u, normalizeStyle as d, onBeforeUnmount as f, onMounted as p, openBlock as m, ref as h, renderList as g, renderSlot as _, toDisplayString as v, unref as y, vModelText as b, watch as x, withDirectives as S, withKeys as C, withModifiers as w } from "vue";
import { getHyperlinkText as T, getRunText as E, isHeadingStyle as D, parseHeadingLevel as O } from "@eigenpal/docx-editor-core/headless";
var k = {
	_lang: "en",
	agentPanel: {
		defaultTitle: "Assistant",
		close: "Close panel",
		resizeHandle: "Resize agent panel",
		thinking: "Assistant is thinking",
		composerPlaceholder: "Ask the assistant…",
		send: "Send",
		timeline: {
			working: "Working… {count, plural, one {# step} other {# steps}}",
			summary: "{count, plural, one {# step} other {# steps}}",
			earlier: "+ {count, plural, one {# earlier step} other {# earlier steps}}"
		}
	},
	aiActions: {
		header: "AI Actions",
		askAi: "Ask AI",
		rewrite: "Rewrite",
		expand: "Expand",
		summarize: "Summarize",
		translate: "Translate",
		explain: "Explain",
		fixGrammar: "Fix grammar",
		makeFormal: "Make formal",
		makeCasual: "Make casual",
		custom: "Custom",
		customPlaceholder: "Custom prompt…"
	},
	aiPreview: {
		defaultTitle: "AI Response",
		labels: {
			rewrite: "Rewrite",
			expand: "Expand",
			summarize: "Summary",
			translate: "Translation",
			explain: "Explanation",
			fixGrammar: "Grammar Fix",
			makeFormal: "Formal Version",
			makeCasual: "Casual Version",
			custom: "AI Response",
			askAi: "AI Response"
		},
		loading: "Processing…",
		original: "Original:",
		suggested: "Suggested:",
		edit: "Edit",
		discard: "Discard",
		accept: "Accept",
		retry: "Retry",
		close: "Close"
	}
}, A = [
	"aria-hidden",
	"aria-label",
	"data-state"
], ee = [
	"aria-label",
	"aria-valuenow",
	"aria-valuemin",
	"aria-valuemax",
	"aria-valuetext"
], te = ["aria-label", "title"], j = "eigenpal:docx-editor:agentPanelWidth", M = /* @__PURE__ */ c({
	__name: "AgentPanel",
	props: {
		title: { default: () => k.agentPanel.defaultTitle },
		closeLabel: { default: () => k.agentPanel.close },
		resizeHandleLabel: { default: () => k.agentPanel.resizeHandle },
		width: { default: void 0 },
		defaultWidth: { default: 360 },
		minWidth: { default: 280 },
		maxWidth: { default: 600 },
		closed: {
			type: Boolean,
			default: !1
		},
		closable: {
			type: Boolean,
			default: !0
		},
		className: { default: "" }
	},
	emits: ["close", "update:width"],
	setup(e, { emit: t }) {
		let r = e, s = t, c = n(() => r.width !== void 0);
		function l() {
			if (typeof window > "u") return r.defaultWidth;
			try {
				let e = window.localStorage.getItem(j);
				if (e) {
					let t = Number(e);
					if (Number.isFinite(t) && t >= r.minWidth && t <= r.maxWidth) return t;
				}
			} catch {}
			return r.defaultWidth;
		}
		let p = h(c.value ? r.width : l()), g = n(() => c.value ? r.width : p.value), y = h(!1), b = null;
		x(() => r.closed, () => {
			b !== null && window.clearTimeout(b), y.value = !0, b = window.setTimeout(() => {
				y.value = !1, b = null;
			}, 260);
		}, { flush: "post" });
		let S = null;
		function w(e) {
			if (!S) return;
			let t = S.startX - e.clientX, n = Math.min(r.maxWidth, Math.max(r.minWidth, S.startWidth + t));
			S.lastWidth = n, c.value || (p.value = n), s("update:width", n);
		}
		function T() {
			if (!S) return;
			let e = S.lastWidth;
			if (S = null, document.removeEventListener("pointermove", w), document.removeEventListener("pointerup", T), !c.value) try {
				window.localStorage.setItem(j, String(e));
			} catch {}
		}
		function E(e) {
			e.preventDefault(), S = {
				startX: e.clientX,
				startWidth: g.value,
				lastWidth: g.value
			}, document.addEventListener("pointermove", w), document.addEventListener("pointerup", T);
		}
		function D(e) {
			r.closable && (e.stopPropagation(), s("close"));
		}
		function O(e) {
			let t = 0, n = null, i = e.shiftKey ? 64 : 16;
			switch (e.key) {
				case "ArrowLeft":
					t = i;
					break;
				case "ArrowRight":
					t = -i;
					break;
				case "Home":
					n = r.maxWidth;
					break;
				case "End":
					n = r.minWidth;
					break;
				default: return;
			}
			e.preventDefault();
			let a = n === null ? Math.min(r.maxWidth, Math.max(r.minWidth, g.value + t)) : n;
			if (c.value || (p.value = a), s("update:width", a), !c.value) try {
				window.localStorage.setItem(j, String(a));
			} catch {}
		}
		f(() => {
			document.removeEventListener("pointermove", w), document.removeEventListener("pointerup", T), b !== null && window.clearTimeout(b);
		});
		let k = n(() => ({
			width: r.closed ? "0px" : `${g.value}px`,
			flex: r.closed ? "0 0 0px" : `0 0 ${g.value}px`,
			height: "calc(100% - 16px)",
			margin: r.closed ? "8px 0 8px 0" : "8px 8px 8px 12px",
			display: "flex",
			flexDirection: "column",
			background: "#ffffff",
			border: r.closed ? "1px solid transparent" : "1px solid #e3e3e3",
			borderRadius: "16px",
			boxShadow: r.closed ? "none" : "0 1px 2px rgba(60,64,67,0.05), 0 4px 12px rgba(60,64,67,0.08)",
			opacity: +!r.closed,
			pointerEvents: r.closed ? "none" : "auto",
			position: "relative",
			boxSizing: "border-box",
			minWidth: r.closed ? 0 : `${r.minWidth}px`,
			overflow: "hidden",
			fontFamily: "'Google Sans', 'Google Sans Text', system-ui, -apple-system, sans-serif",
			transition: y.value ? "flex-basis 220ms cubic-bezier(0.4, 0, 0.2, 1), width 220ms cubic-bezier(0.4, 0, 0.2, 1), margin 220ms cubic-bezier(0.4, 0, 0.2, 1), opacity 180ms ease, box-shadow 220ms ease, border-color 220ms ease" : "opacity 180ms ease, box-shadow 220ms ease, border-color 220ms ease"
		})), M = {
			position: "absolute",
			left: "-3px",
			top: 0,
			bottom: 0,
			width: "6px",
			cursor: "col-resize",
			touchAction: "none",
			zIndex: 1
		}, N = {
			display: "flex",
			alignItems: "center",
			gap: "10px",
			padding: "14px 16px 10px",
			flex: "0 0 auto",
			background: "#ffffff"
		}, P = {
			display: "inline-flex",
			alignItems: "center",
			color: "#0b57d0"
		}, F = {
			flex: 1,
			fontSize: "15px",
			fontWeight: 500,
			color: "#1f1f1f",
			letterSpacing: "0.1px",
			overflow: "hidden",
			textOverflow: "ellipsis",
			whiteSpace: "nowrap"
		}, I = {
			border: "none",
			background: "transparent",
			padding: "6px",
			cursor: "pointer",
			display: "inline-flex",
			alignItems: "center",
			justifyContent: "center",
			color: "#444746",
			borderRadius: "999px",
			transition: "background 0.15s"
		}, ne = {
			flex: 1,
			minHeight: 0,
			overflow: "hidden",
			display: "flex",
			flexDirection: "column"
		};
		function L(e, t) {
			let n = t.currentTarget;
			n && (n.style.background = e ? "#f1f3f4" : "transparent");
		}
		return (t, n) => (m(), a("div", {
			class: u(["ep-agent-panel", e.className]),
			style: d(k.value),
			"aria-hidden": e.closed,
			"aria-label": e.title,
			role: "complementary",
			"data-testid": "agent-panel",
			"data-state": e.closed ? "closed" : "open",
			onKeydown: C(D, ["esc"])
		}, [
			o("div", {
				role: "separator",
				"aria-orientation": "vertical",
				"aria-label": e.resizeHandleLabel,
				"aria-valuenow": g.value,
				"aria-valuemin": e.minWidth,
				"aria-valuemax": e.maxWidth,
				"aria-valuetext": `${g.value} pixels wide`,
				tabindex: "0",
				style: M,
				"data-testid": "agent-panel-resize-handle",
				onPointerdown: E,
				onKeydown: O
			}, null, 40, ee),
			o("div", { style: N }, [
				o("span", { style: P }, [_(t.$slots, "icon", {}, () => [n[3] ||= o("svg", {
					viewBox: "0 -960 960 960",
					width: "22",
					height: "22",
					fill: "currentColor",
					"aria-hidden": "true"
				}, [o("path", { d: "m760-600-50-110-110-50 110-50 50-110 50 110 110 50-110 50-50 110Zm0 560-50-110-110-50 110-50 50-110 50 110 110 50-110 50-50 110ZM360-160 260-380 40-480l220-100 100-220 100 220 220 100-220 100-100 220Zm0-194 40-86 86-40-86-40-40-86-40 86-86 40 86 40 40 86Zm0-126Z" })], -1)])]),
				o("span", { style: F }, v(e.title), 1),
				e.closable ? (m(), a("button", {
					key: 0,
					type: "button",
					"aria-label": e.closeLabel,
					title: e.closeLabel,
					"data-testid": "agent-panel-close",
					style: I,
					onClick: n[0] ||= (e) => s("close"),
					onMouseenter: n[1] ||= (e) => L(!0, e),
					onMouseleave: n[2] ||= (e) => L(!1, e)
				}, [...n[4] ||= [o("svg", {
					viewBox: "0 -960 960 960",
					width: "18",
					height: "18",
					fill: "currentColor",
					"aria-hidden": "true"
				}, [o("path", { d: "m256-200-56-56 224-224-224-224 56-56 224 224 224-224 56 56-224 224 224 224-56 56-224-224-224 224Z" })], -1)]], 40, te)) : i("", !0)
			]),
			o("div", { style: ne }, [_(t.$slots, "default")])
		], 46, A));
	}
});
//#endregion
//#region src/i18n/format-message.ts
function N(e) {
	let t = {}, n = /(=\d+|\w+)\s*\{([^}]*)\}/g, r;
	for (; (r = n.exec(e)) !== null;) t[r[1]] = r[2];
	return t;
}
function P(e, t, n = "en") {
	return t ? e.replace(/\{(\w+),\s*plural,\s*((?:[^{}]|\{[^{}]*\})*)\}/g, (e, r, i) => {
		let a = Number(t[r]);
		if (isNaN(a)) return e;
		let o = N(i), s = o[`=${a}`];
		if (s !== void 0) return s.replace(/#/g, String(a));
		let c;
		try {
			c = new Intl.PluralRules(n).select(a);
		} catch {
			c = a === 1 ? "one" : "other";
		}
		return (o[c] ?? o.other ?? "").replace(/#/g, String(a));
	}).replace(/\{(\w+)\}/g, (e, n) => {
		let r = t[n];
		return r === void 0 ? `{${n}}` : String(r);
	}) : e;
}
//#endregion
//#region src/agent-types.ts
function F(e) {
	let t = e.replace(/_/g, " ");
	return t.charAt(0).toUpperCase() + t.slice(1);
}
//#endregion
//#region src/vue/components/AgentTimeline.vue?vue&type=script&setup=true&lang.ts
var I = ["aria-expanded"], ne = {
	key: 1,
	style: {
		marginRight: "6px",
		display: "inline-flex"
	},
	"aria-hidden": "true"
}, L = /* @__PURE__ */ c({
	__name: "AgentTimeline",
	props: {
		toolCalls: {},
		streaming: {
			type: Boolean,
			default: !1
		},
		maxVisibleCalls: { default: 3 },
		humanizeName: {
			type: Function,
			default: void 0
		},
		workingLabel: {
			type: Function,
			default: void 0
		},
		summaryLabel: {
			type: Function,
			default: void 0
		},
		earlierLabel: {
			type: Function,
			default: void 0
		}
	},
	setup(t) {
		let r = t, c = h(null), l = n(() => c.value === null ? !!r.streaming : c.value), u = (e) => (r.humanizeName ?? F)(e), f = (e) => (r.earlierLabel ?? ((e) => P(k.agentPanel.timeline.earlier, { count: e })))(e), p = n(() => {
			let e = r.toolCalls.length;
			return r.streaming ? (r.workingLabel ?? ((e) => P(k.agentPanel.timeline.working, { count: e })))(e) : (r.summaryLabel ?? ((e) => P(k.agentPanel.timeline.summary, { count: e })))(e);
		}), _ = n(() => r.toolCalls.slice(-r.maxVisibleCalls)), y = n(() => Math.max(0, r.toolCalls.length - _.value.length)), b = n(() => ({
			fontSize: "12px",
			color: "#5f6368",
			transition: "transform 0.15s ease",
			marginLeft: "8px",
			display: "inline-block",
			transform: l.value ? "rotate(180deg)" : "rotate(0deg)"
		}));
		function x(e = !1) {
			return {
				marginRight: e ? 0 : "6px",
				display: "inline-flex",
				animation: "epAgentSpin 0.8s linear infinite",
				flexShrink: 0
			};
		}
		let S = {
			timelineWrap: {
				alignSelf: "flex-start",
				maxWidth: "92%",
				background: "#fff",
				border: "1px solid #e1e3e6",
				borderRadius: "12px",
				overflow: "hidden",
				fontFamily: "'Google Sans Text', system-ui, sans-serif"
			},
			timelineHeader: {
				display: "flex",
				alignItems: "center",
				justifyContent: "space-between",
				width: "100%",
				padding: "8px 12px",
				background: "transparent",
				border: "none",
				cursor: "pointer",
				fontSize: "12.5px",
				color: "#1f1f1f",
				fontFamily: "inherit"
			},
			timelineSummary: {
				display: "inline-flex",
				alignItems: "center",
				fontWeight: 500
			},
			timelineList: {
				listStyle: "none",
				margin: 0,
				padding: "4px 12px 10px 12px",
				display: "flex",
				flexDirection: "column",
				gap: "6px",
				borderTop: "1px solid #ececf0"
			},
			timelineItem: {
				display: "flex",
				alignItems: "center",
				gap: "8px",
				fontSize: "12px",
				color: "#444746"
			},
			timelineDot: {
				width: "6px",
				height: "6px",
				borderRadius: "50%",
				flexShrink: 0
			},
			timelineCall: {
				display: "inline-flex",
				flexDirection: "column",
				minWidth: 0
			},
			timelineCallName: { color: "#1f1f1f" },
			timelineError: {
				color: "#d93025",
				fontSize: "11px"
			},
			timelineMore: {
				fontSize: "11px",
				color: "#5f6368",
				fontStyle: "italic",
				paddingLeft: "14px"
			}
		};
		return (n, r) => t.toolCalls.length > 0 ? (m(), a("div", {
			key: 0,
			style: d(S.timelineWrap),
			"data-testid": "agent-timeline"
		}, [o("button", {
			type: "button",
			"aria-expanded": l.value,
			style: d(S.timelineHeader),
			"data-testid": "agent-timeline-toggle",
			onClick: r[0] ||= (e) => c.value = !l.value
		}, [o("span", { style: d(S.timelineSummary) }, [t.streaming ? (m(), a("span", {
			key: 0,
			style: d(x()),
			"aria-hidden": "true"
		}, [...r[1] ||= [o("svg", {
			width: "12",
			height: "12",
			viewBox: "0 0 24 24",
			fill: "none"
		}, [o("circle", {
			cx: "12",
			cy: "12",
			r: "9",
			stroke: "#dadce0",
			"stroke-width": "3",
			fill: "none"
		}), o("path", {
			d: "M21 12a9 9 0 0 0-9-9",
			stroke: "#0b57d0",
			"stroke-width": "3",
			"stroke-linecap": "round"
		})], -1)]], 4)) : (m(), a("span", ne, [...r[2] ||= [o("svg", {
			width: "12",
			height: "12",
			viewBox: "0 0 24 24",
			fill: "none"
		}, [o("path", {
			d: "M5 13l4 4L19 7",
			stroke: "#1e8e3e",
			"stroke-width": "2.5",
			"stroke-linecap": "round",
			"stroke-linejoin": "round"
		})], -1)]])), s(" " + v(p.value), 1)], 4), o("span", {
			style: d(b.value),
			"aria-hidden": "true"
		}, "▾", 4)], 12, I), l.value ? (m(), a("ol", {
			key: 0,
			style: d(S.timelineList)
		}, [y.value > 0 ? (m(), a("li", {
			key: 0,
			style: d(S.timelineMore),
			"data-testid": "agent-timeline-earlier"
		}, v(f(y.value)), 5)) : i("", !0), (m(!0), a(e, null, g(_.value, (e) => (m(), a("li", {
			key: e.id,
			style: d(S.timelineItem)
		}, [e.status === "running" ? (m(), a("span", {
			key: 0,
			style: d(x(!0)),
			"aria-hidden": "true"
		}, [...r[3] ||= [o("svg", {
			width: "10",
			height: "10",
			viewBox: "0 0 24 24",
			fill: "none"
		}, [o("circle", {
			cx: "12",
			cy: "12",
			r: "9",
			stroke: "#dadce0",
			"stroke-width": "3",
			fill: "none"
		}), o("path", {
			d: "M21 12a9 9 0 0 0-9-9",
			stroke: "#0b57d0",
			"stroke-width": "3",
			"stroke-linecap": "round"
		})], -1)]], 4)) : (m(), a("span", {
			key: 1,
			style: d({
				...S.timelineDot,
				background: e.status === "error" ? "#d93025" : "#1e8e3e"
			}),
			"aria-hidden": "true"
		}, null, 4)), o("span", { style: d(S.timelineCall) }, [o("span", { style: d(S.timelineCallName) }, v(u(e.name)), 5), e.error ? (m(), a("span", {
			key: 0,
			style: d(S.timelineError)
		}, v(e.error), 5)) : i("", !0)], 4)], 4))), 128))], 4)) : i("", !0)], 4)) : i("", !0);
	}
}), re = ["data-role"], ie = ["aria-label"], R = "ep-agent-chat-keyframes", ae = /* @__PURE__ */ c({
	__name: "AgentChatLog",
	props: {
		messages: {},
		loading: {
			type: Boolean,
			default: !1
		},
		error: { default: null },
		thinkingLabel: { default: () => k.agentPanel.thinking },
		workingLabel: {
			type: Function,
			default: void 0
		},
		summaryLabel: {
			type: Function,
			default: void 0
		},
		earlierLabel: {
			type: Function,
			default: void 0
		},
		autoScroll: {
			type: Boolean,
			default: !0
		},
		humanizeToolName: {
			type: Function,
			default: void 0
		},
		maxVisibleCalls: { default: 3 },
		className: { default: "" }
	},
	setup(t) {
		let s = t, c = h(null), f = n(() => s.messages.length === 0 && !s.loading && !s.error);
		function y() {
			if (!s.autoScroll) return;
			let e = s.messages[s.messages.length - 1]?.status === "streaming";
			c.value?.scrollIntoView({
				behavior: e || s.loading ? "auto" : "smooth",
				block: "end"
			});
		}
		p(() => {
			l(y), b();
		}), x([
			() => s.messages.length,
			() => s.messages[s.messages.length - 1]?.toolCalls?.length ?? 0,
			() => s.loading,
			() => s.messages[s.messages.length - 1]?.status === "streaming"
		], y, { flush: "post" });
		function b() {
			if (typeof document > "u" || document.getElementById(R)) return;
			let e = document.createElement("style");
			e.id = R, e.textContent = "\n@keyframes epAgentDot {\n  0%, 80%, 100% { transform: scale(0.6); opacity: 0.4; }\n  40% { transform: scale(1); opacity: 1; }\n}\n@keyframes epAgentSpin {\n  to { transform: rotate(360deg); }\n}\n", document.head.appendChild(e);
		}
		let S = {
			flex: 1,
			overflow: "auto",
			padding: "16px 14px 8px",
			display: "flex",
			flexDirection: "column",
			gap: "10px"
		}, C = {
			messageGroup: {
				display: "flex",
				flexDirection: "column",
				gap: "6px",
				width: "100%"
			},
			userBubble: {
				background: "#0b57d0",
				color: "#fff",
				padding: "10px 14px",
				borderRadius: "20px 20px 4px 20px",
				fontSize: "13.5px",
				lineHeight: 1.5,
				alignSelf: "flex-end",
				maxWidth: "88%",
				whiteSpace: "pre-wrap",
				wordBreak: "break-word",
				fontFamily: "'Google Sans', 'Google Sans Text', system-ui, -apple-system, sans-serif"
			},
			assistantBubble: {
				background: "#f0f4f9",
				color: "#1f1f1f",
				padding: "12px 16px",
				borderRadius: "20px 20px 20px 4px",
				fontSize: "13.5px",
				lineHeight: 1.55,
				alignSelf: "flex-start",
				maxWidth: "92%",
				whiteSpace: "pre-wrap",
				wordBreak: "break-word",
				fontFamily: "'Google Sans', 'Google Sans Text', system-ui, -apple-system, sans-serif"
			},
			thinkingBubble: {
				background: "#f0f4f9",
				padding: "12px 16px",
				borderRadius: "20px 20px 20px 4px",
				alignSelf: "flex-start",
				display: "flex",
				gap: "4px",
				alignItems: "center"
			},
			dot: {
				width: "6px",
				height: "6px",
				borderRadius: "50%",
				background: "#5f6368",
				display: "inline-block",
				animation: "epAgentDot 1.4s infinite ease-in-out"
			},
			errorBubble: {
				background: "#fce8e6",
				color: "#b3261e",
				padding: "10px 14px",
				borderRadius: "16px",
				fontSize: "12.5px",
				alignSelf: "flex-start",
				maxWidth: "92%",
				whiteSpace: "pre-wrap",
				fontFamily: "'Google Sans Text', system-ui, sans-serif"
			}
		};
		return (n, s) => (m(), a("div", {
			class: u(["ep-agent-chat-log", t.className]),
			style: S,
			role: "log",
			"aria-live": "polite",
			"aria-atomic": "false"
		}, [
			f.value ? _(n.$slots, "empty", { key: 0 }) : i("", !0),
			(m(!0), a(e, null, g(t.messages, (e) => (m(), a("div", {
				key: e.id,
				style: d(C.messageGroup),
				"data-role": e.role
			}, [e.role === "assistant" && e.toolCalls && e.toolCalls.length > 0 ? (m(), r(L, {
				key: 0,
				"tool-calls": e.toolCalls,
				streaming: e.status === "streaming",
				"humanize-name": t.humanizeToolName,
				"max-visible-calls": t.maxVisibleCalls,
				"working-label": t.workingLabel,
				"summary-label": t.summaryLabel,
				"earlier-label": t.earlierLabel
			}, null, 8, [
				"tool-calls",
				"streaming",
				"humanize-name",
				"max-visible-calls",
				"working-label",
				"summary-label",
				"earlier-label"
			])) : i("", !0), e.text.length > 0 ? (m(), a("div", {
				key: 1,
				style: d(e.role === "user" ? C.userBubble : C.assistantBubble)
			}, v(e.text), 5)) : i("", !0)], 12, re))), 128)),
			t.loading ? (m(), a("div", {
				key: 1,
				style: d(C.thinkingBubble),
				"aria-label": t.thinkingLabel
			}, [
				o("span", { style: d({
					...C.dot,
					animationDelay: "0s"
				}) }, null, 4),
				o("span", { style: d({
					...C.dot,
					animationDelay: "0.15s"
				}) }, null, 4),
				o("span", { style: d({
					...C.dot,
					animationDelay: "0.3s"
				}) }, null, 4)
			], 12, ie)) : i("", !0),
			t.error ? (m(), a("div", {
				key: 2,
				style: d(C.errorBubble),
				role: "alert"
			}, v(t.error), 5)) : i("", !0),
			o("div", {
				ref_key: "endEl",
				ref: c
			}, null, 512)
		], 2));
	}
}), oe = [
	"value",
	"placeholder",
	"disabled"
], se = ["aria-label", "disabled"], ce = /* @__PURE__ */ c({
	__name: "AgentComposer",
	props: {
		modelValue: {},
		disabled: {
			type: Boolean,
			default: !1
		},
		placeholder: { default: () => k.agentPanel.composerPlaceholder },
		sendLabel: { default: () => k.agentPanel.send },
		className: { default: "" }
	},
	emits: ["update:modelValue", "submit"],
	setup(e, { emit: t }) {
		let r = e, s = t, c = n(() => r.modelValue.trim().length > 0 && !r.disabled);
		function l() {
			c.value && s("submit");
		}
		let f = n(() => ({
			width: "36px",
			height: "36px",
			borderRadius: "50%",
			border: "none",
			background: "#0b57d0",
			color: "#fff",
			flexShrink: 0,
			display: "inline-flex",
			alignItems: "center",
			justifyContent: "center",
			transition: "background 0.15s, opacity 0.15s, transform 0.15s",
			opacity: c.value ? 1 : .35,
			cursor: c.value ? "pointer" : "not-allowed"
		})), p = {
			composerWrap: {
				padding: "8px 12px 14px",
				background: "#fff",
				flex: "0 0 auto"
			},
			composerShell: {
				display: "flex",
				alignItems: "center",
				gap: "4px",
				padding: "6px 6px 6px 18px",
				background: "#fff",
				border: "1px solid #c4c7c5",
				borderRadius: "28px",
				boxShadow: "0 1px 2px rgba(60,64,67,0.04)",
				transition: "border-color 0.15s, box-shadow 0.15s",
				fontFamily: "'Google Sans Text', system-ui, sans-serif"
			},
			composerInput: {
				flex: 1,
				padding: "8px 0",
				fontSize: "14px",
				border: "none",
				outline: "none",
				background: "transparent",
				fontFamily: "inherit",
				color: "#1f1f1f"
			},
			footnote: {
				fontSize: "11px",
				color: "#5f6368",
				textAlign: "center",
				marginTop: "10px",
				fontFamily: "'Google Sans Text', system-ui, sans-serif"
			}
		};
		return (t, n) => (m(), a("form", {
			class: u(["ep-agent-composer", e.className]),
			style: d(p.composerWrap),
			onSubmit: w(l, ["prevent"])
		}, [o("div", { style: d(p.composerShell) }, [o("input", {
			style: d(p.composerInput),
			value: e.modelValue,
			placeholder: e.placeholder,
			disabled: e.disabled,
			onInput: n[0] ||= (e) => s("update:modelValue", e.target.value)
		}, null, 44, oe), o("button", {
			type: "submit",
			"aria-label": e.sendLabel,
			disabled: !c.value,
			style: d(f.value)
		}, [...n[1] ||= [o("svg", {
			width: "16",
			height: "16",
			viewBox: "0 0 24 24",
			fill: "none",
			"aria-hidden": "true"
		}, [o("path", {
			d: "M12 19V5M5 12l7-7 7 7",
			stroke: "currentColor",
			"stroke-width": "2",
			"stroke-linecap": "round",
			"stroke-linejoin": "round"
		})], -1)]], 12, se)], 4), t.$slots.footnote ? (m(), a("div", {
			key: 0,
			style: d(p.footnote)
		}, [_(t.$slots, "footnote")], 4)) : i("", !0)], 38));
	}
}), le = ["disabled"], ue = /* @__PURE__ */ c({
	__name: "AgentSuggestionChip",
	props: {
		label: {},
		disabled: { type: Boolean }
	},
	emits: ["click"],
	setup(e, { emit: t }) {
		let n = t, r = {
			textAlign: "left",
			background: "#fff",
			border: "1px solid #dadce0",
			borderRadius: "18px",
			padding: "10px 14px",
			fontSize: "13px",
			color: "#1f1f1f",
			cursor: "pointer",
			fontFamily: "'Google Sans Text', system-ui, sans-serif",
			lineHeight: 1.4,
			transition: "background 0.15s, border-color 0.15s, box-shadow 0.15s",
			width: "100%"
		};
		return (t, i) => (m(), a("button", {
			type: "button",
			style: r,
			disabled: e.disabled,
			onClick: i[0] ||= (e) => n("click")
		}, v(e.label), 9, le));
	}
}), de = ["aria-label"], fe = {
	class: "ai-ctx-menu__header",
	"aria-hidden": "true"
}, pe = ["onMousedown"], me = ["innerHTML"], he = { class: "ai-ctx-menu__label" }, ge = {
	key: 0,
	class: "ai-ctx-menu__custom"
}, _e = ["placeholder"], ve = /* @__PURE__ */ c({
	__name: "AIContextMenu",
	props: {
		isOpen: { type: Boolean },
		position: {},
		selectedText: {},
		showCustomPrompt: {
			type: Boolean,
			default: !0
		},
		labels: { default: () => ({}) }
	},
	emits: ["close", "action"],
	setup(s, { emit: c }) {
		let u = s, f = c, p = h(""), _ = h([]);
		x(() => u.isOpen, (e) => {
			e && l(() => _.value[0]?.focus());
		}, { immediate: !0 });
		let y = [
			{
				id: "rewrite",
				icon: "&#x270D;"
			},
			{
				id: "expand",
				icon: "&#x2194;"
			},
			{
				id: "summarize",
				icon: "&#x1F4DD;"
			},
			{
				id: "fixGrammar",
				icon: "&#x2714;"
			},
			{
				id: "makeFormal",
				icon: "&#x1F454;"
			},
			{
				id: "makeCasual",
				icon: "&#x1F60A;"
			},
			{
				id: "translate",
				icon: "&#x1F310;"
			},
			{
				id: "explain",
				icon: "&#x1F4A1;"
			}
		], T = n(() => u.labels?.header ?? k.aiActions.header), E = n(() => u.labels?.customPlaceholder ?? k.aiActions.customPlaceholder);
		function D(e) {
			return u.labels?.[e] || (k.aiActions[e] ?? e);
		}
		let O = n(() => {
			let e = u.position.x, t = u.position.y;
			return typeof window < "u" && (e + 220 + 10 > window.innerWidth && (e = window.innerWidth - 220 - 10), t + 360 + 10 > window.innerHeight && (t = window.innerHeight - 360 - 10)), {
				position: "fixed",
				left: `${e}px`,
				top: `${t}px`,
				zIndex: 500
			};
		});
		function A(e) {
			f("action", e, e === "custom" ? p.value : void 0), p.value = "", f("close");
		}
		return (n, c) => (m(), r(t, { to: "body" }, [s.isOpen ? (m(), a("div", {
			key: 0,
			class: "ai-ctx-backdrop",
			onMousedown: c[0] ||= (e) => f("close"),
			onContextmenu: c[1] ||= w((e) => f("close"), ["prevent"])
		}, null, 32)) : i("", !0), s.isOpen ? (m(), a("div", {
			key: 1,
			class: "ai-ctx-menu",
			"aria-label": T.value,
			style: d(O.value),
			tabindex: "-1",
			onContextmenu: c[4] ||= w(() => {}, ["prevent"]),
			onKeydown: c[5] ||= C(w((e) => f("close"), ["stop"]), ["esc"])
		}, [
			o("div", fe, v(T.value), 1),
			(m(), a(e, null, g(y, (e) => o("button", {
				key: e.id,
				ref_for: !0,
				ref_key: "itemRefs",
				ref: _,
				class: "ai-ctx-menu__item",
				onMousedown: w((t) => A(e.id), ["prevent"])
			}, [o("span", {
				class: "ai-ctx-menu__icon",
				innerHTML: e.icon
			}, null, 8, me), o("span", he, v(D(e.id)), 1)], 40, pe)), 64)),
			c[6] ||= o("div", { class: "ai-ctx-menu__divider" }, null, -1),
			s.showCustomPrompt ? (m(), a("div", ge, [S(o("input", {
				"onUpdate:modelValue": c[2] ||= (e) => p.value = e,
				class: "ai-ctx-menu__input",
				placeholder: E.value,
				onKeydown: c[3] ||= C(w((e) => A("custom"), ["prevent"]), ["enter"])
			}, null, 40, _e), [[b, p.value]])])) : i("", !0)
		], 44, de)) : i("", !0)]));
	}
}), z = (e, t) => {
	let n = e.__vccOpts || e;
	for (let [e, r] of t) n[e] = r;
	return n;
}, ye = /* @__PURE__ */ z(ve, [["__scopeId", "data-v-c30da7a7"]]), be = ["aria-label"], xe = { class: "ai-preview__header" }, Se = { class: "ai-preview__title" }, Ce = ["aria-label"], we = {
	key: 0,
	class: "ai-preview__loading"
}, Te = {
	key: 1,
	class: "ai-preview__error"
}, Ee = {
	key: 2,
	class: "ai-preview__content"
}, De = {
	key: 0,
	class: "ai-preview__diff"
}, Oe = { class: "ai-preview__diff-label" }, ke = { class: "ai-preview__diff-text ai-preview__diff-text--old" }, Ae = { class: "ai-preview__diff-label" }, je = { class: "ai-preview__diff-text ai-preview__diff-text--new" }, Me = {
	key: 1,
	class: "ai-preview__result"
}, Ne = {
	key: 3,
	class: "ai-preview__footer"
}, Pe = /* @__PURE__ */ z(/* @__PURE__ */ c({
	__name: "AIResponsePreview",
	props: {
		isVisible: { type: Boolean },
		originalText: {},
		responseText: {},
		action: {},
		isLoading: { type: Boolean },
		error: { default: void 0 },
		allowEdit: {
			type: Boolean,
			default: !0
		},
		showDiff: {
			type: Boolean,
			default: !0
		},
		showRetry: {
			type: Boolean,
			default: !0
		},
		labels: { default: () => ({}) }
	},
	emits: [
		"accept",
		"reject",
		"retry"
	],
	setup(e, { emit: t }) {
		let r = e, s = t, c = h(!1), l = h(""), u = n(() => ({
			loading: r.labels?.loading ?? k.aiPreview.loading,
			original: r.labels?.original ?? k.aiPreview.original,
			suggested: r.labels?.suggested ?? k.aiPreview.suggested,
			edit: r.labels?.edit ?? k.aiPreview.edit,
			discard: r.labels?.discard ?? k.aiPreview.discard,
			accept: r.labels?.accept ?? k.aiPreview.accept,
			retry: r.labels?.retry ?? k.aiPreview.retry,
			close: r.labels?.close ?? k.aiPreview.close
		})), d = n(() => r.labels?.actionTitles?.[r.action] || (k.aiPreview.labels[r.action] ?? k.aiPreview.defaultTitle));
		x(() => r.responseText, (e) => {
			l.value = e, c.value = !1;
		});
		function f() {
			s("accept", c.value ? l.value : r.responseText);
		}
		return (t, n) => e.isVisible ? (m(), a("div", {
			key: 0,
			class: "ai-preview",
			role: "region",
			"aria-label": d.value,
			tabindex: "-1",
			onKeydown: n[5] ||= C(w((e) => s("reject"), ["stop"]), ["esc"])
		}, [
			o("div", xe, [o("span", Se, v(d.value), 1), o("button", {
				class: "ai-preview__close",
				"aria-label": u.value.close,
				onClick: n[0] ||= (e) => s("reject")
			}, " ✕ ", 8, Ce)]),
			e.isLoading ? (m(), a("div", we, [n[6] ||= o("span", { class: "ai-preview__spinner" }, null, -1), o("span", null, v(u.value.loading), 1)])) : e.error ? (m(), a("div", Te, [o("span", null, v(e.error), 1), e.showRetry ? (m(), a("button", {
				key: 0,
				class: "ai-preview__retry",
				onMousedown: n[1] ||= w((e) => s("retry"), ["prevent"])
			}, v(u.value.retry), 33)) : i("", !0)])) : (m(), a("div", Ee, [e.showDiff ? (m(), a("div", De, [
				o("div", Oe, v(u.value.original), 1),
				o("div", ke, v(e.originalText), 1),
				o("div", Ae, v(u.value.suggested), 1),
				o("div", je, v(e.responseText), 1)
			])) : (m(), a("div", Me, v(e.responseText), 1)), e.allowEdit && c.value ? S((m(), a("textarea", {
				key: 2,
				"onUpdate:modelValue": n[2] ||= (e) => l.value = e,
				class: "ai-preview__textarea",
				rows: "4"
			}, null, 512)), [[b, l.value]]) : i("", !0)])),
			!e.isLoading && !e.error ? (m(), a("div", Ne, [
				e.allowEdit && !c.value ? (m(), a("button", {
					key: 0,
					class: "ai-preview__btn",
					onMousedown: n[3] ||= w((e) => c.value = !0, ["prevent"])
				}, v(u.value.edit), 33)) : i("", !0),
				o("button", {
					class: "ai-preview__btn",
					onMousedown: n[4] ||= w((e) => s("reject"), ["prevent"])
				}, v(u.value.discard), 33),
				o("button", {
					class: "ai-preview__btn ai-preview__btn--primary",
					onMousedown: w(f, ["prevent"])
				}, v(u.value.accept), 33)
			])) : i("", !0)
		], 40, be)) : i("", !0);
	}
}), [["__scopeId", "data-v-1226b741"]]), B = [
	"single",
	"words",
	"double",
	"thick",
	"dotted",
	"dottedHeavy",
	"dash",
	"dashedHeavy",
	"dashLong",
	"dashLongHeavy",
	"dotDash",
	"dashDotHeavy",
	"dotDotDash",
	"dashDotDotHeavy",
	"wave",
	"wavyHeavy",
	"wavyDouble",
	"none"
], V = [
	"black",
	"blue",
	"cyan",
	"darkBlue",
	"darkCyan",
	"darkGray",
	"darkGreen",
	"darkMagenta",
	"darkRed",
	"darkYellow",
	"green",
	"lightGray",
	"magenta",
	"red",
	"white",
	"yellow",
	"none"
], Fe = {
	name: "apply_formatting",
	displayName: "Applying formatting",
	description: "Apply character formatting (bold, italic, underline, strike, color, highlight, font size, font family) to a paragraph or to a unique phrase within it. Pass `search` to scope the change to part of the paragraph; omit it to format the whole paragraph. Direct edit — does not create a tracked change. Pass `false` to clear a mark; omit a key to leave it untouched. Color uses `{rgb: \"FF0000\"}` (no hash) or `{themeColor: \"accent1\"}`. Font size is in points. Font family takes `{ascii, hAnsi}`.",
	inputSchema: {
		type: "object",
		properties: {
			paraId: {
				type: "string",
				description: "Paragraph id from read_document / find_text."
			},
			search: {
				type: "string",
				description: "Optional: format only this exact phrase within the paragraph (must be unique)."
			},
			marks: {
				type: "object",
				description: "Marks to set or clear. Omit a key to leave it untouched.",
				properties: {
					bold: { type: "boolean" },
					italic: { type: "boolean" },
					underline: { description: "true → single underline; false → clear; or { style: \"single\"|\"double\"|\"thick\"|\"dotted\"|\"dottedHeavy\"|\"dash\"|\"dashedHeavy\"|\"dashLong\"|\"dashLongHeavy\"|\"dotDash\"|\"dashDotHeavy\"|\"dotDotDash\"|\"dashDotDotHeavy\"|\"wave\"|\"wavyHeavy\"|\"wavyDouble\"|\"words\"|\"none\" }. Other values are rejected." },
					strike: { type: "boolean" },
					color: {
						type: "object",
						description: "Either {rgb: \"RRGGBB\"} (no hash) or {themeColor: \"accent1\"|\"text1\"|...}.",
						properties: {
							rgb: { type: "string" },
							themeColor: { type: "string" }
						}
					},
					highlight: {
						type: "string",
						enum: [...V],
						description: "Highlight color — must be one of the Word-supported names: " + V.join(", ") + ". Pass \"none\" to clear. Hex values are rejected (Word does not accept hex for <w:highlight>)."
					},
					fontSize: {
						type: "number",
						description: "Size in points (e.g. 12, 14, 24)."
					},
					fontFamily: {
						type: "object",
						properties: {
							ascii: { type: "string" },
							hAnsi: { type: "string" }
						}
					}
				}
			}
		},
		required: ["paraId", "marks"]
	},
	handler: (e, t) => {
		if (!e.marks || Object.keys(e.marks).length === 0) return {
			success: !1,
			error: "No marks provided. Specify at least one of bold/italic/etc."
		};
		let n = typeof e.marks.underline == "object" && e.marks.underline !== null ? e.marks.underline.style : void 0;
		if (n && !B.includes(n)) return {
			success: !1,
			error: `Invalid underline.style "${n}". Must be one of: ${B.join(", ")}.`
		};
		let r = typeof e.marks.highlight == "string" ? e.marks.highlight : void 0;
		if (r && !V.includes(r)) return {
			success: !1,
			error: `Invalid highlight "${r}". Must be one of: ${V.join(", ")}. Hex values are not supported by Word's highlight attribute.`
		};
		let i = r === "none" ? {
			...e.marks,
			highlight: ""
		} : e.marks;
		return t.applyFormatting({
			paraId: e.paraId,
			search: e.search,
			marks: i
		}) ? {
			success: !0,
			data: `Formatting applied to ${e.search ? `"${e.search}" in ${e.paraId}` : e.paraId}.`
		} : {
			success: !1,
			error: "Could not apply formatting. The paraId may not exist, or `search` is missing / ambiguous."
		};
	}
}, Ie = {
	name: "set_paragraph_style",
	displayName: "Setting paragraph style",
	description: "Apply a paragraph style by id (e.g. \"Heading1\", \"Heading2\", \"Title\", \"Quote\", \"Normal\"). The styleId must exist in the document's style definitions — unknown ids are no-ops. Direct edit, not a tracked change.",
	inputSchema: {
		type: "object",
		properties: {
			paraId: {
				type: "string",
				description: "Paragraph id from read_document / find_text."
			},
			styleId: {
				type: "string",
				description: "Style id (e.g. \"Heading1\", \"Title\", \"Quote\", \"Normal\")."
			}
		},
		required: ["paraId", "styleId"]
	},
	handler: (e, t) => t.setParagraphStyle({
		paraId: e.paraId,
		styleId: e.styleId
	}) ? {
		success: !0,
		data: `Style "${e.styleId}" applied to ${e.paraId}.`
	} : {
		success: !1,
		error: `Could not set style. paraId "${e.paraId}" not found, or styleId "${e.styleId}" is not defined.`
	}
}, H = [
	"page",
	"sectionNextPage",
	"sectionContinuous"
], Le = {
	page: "Page break",
	sectionNextPage: "Section break (next page)",
	sectionContinuous: "Section break (continuous)"
}, U = [
	{
		name: "read_document",
		displayName: "Reading document",
		description: "Read the document content. Returns lines tagged with a stable paragraph id, e.g. \"[2A1F3B] First paragraph\". Use the bracketed id as `paraId` when commenting or suggesting changes — it survives edits, unlike ordinal indices. Returns the vanilla document (the doc as it exists right now, before any tracked suggestions are accepted): pending insertions are HIDDEN, pending deletions are shown as plain text (still part of the document until accepted), and comment markers are stripped. Use read_changes / read_comments to inspect what is pending.",
		inputSchema: {
			type: "object",
			properties: {
				fromIndex: {
					type: "number",
					description: "Start ordinal index (inclusive). Optional."
				},
				toIndex: {
					type: "number",
					description: "End ordinal index (inclusive). Optional."
				}
			}
		},
		handler: (e, t) => ({
			success: !0,
			data: t.getContentAsText({
				fromIndex: e.fromIndex,
				toIndex: e.toIndex,
				includeTrackedChanges: !1,
				includeCommentAnchors: !1
			})
		})
	},
	{
		name: "read_selection",
		displayName: "Reading selection",
		description: "Read the user's current cursor or selection. Returns the selected text, the paragraph it lives in, and that paragraph's `paraId`. Use this when the user asks \"fix this\" or \"review what I have selected\".",
		inputSchema: {
			type: "object",
			properties: {}
		},
		handler: (e, t) => {
			let n = t.getSelection();
			return n ? {
				success: !0,
				data: n
			} : {
				success: !1,
				error: "No selection (editor not focused)."
			};
		}
	},
	{
		name: "read_page",
		displayName: "Reading page",
		description: "Read the contents of one rendered page (1-indexed). Returns paragraphs on the page, each tagged with its stable paraId. Use this when the user asks \"summarize page 3\" or \"comment on what's on this page\".",
		inputSchema: {
			type: "object",
			properties: { pageNumber: {
				type: "number",
				description: "1-indexed page number."
			} },
			required: ["pageNumber"]
		},
		handler: (e, t) => {
			let n = t.getPage(e.pageNumber);
			if (!n) {
				let n = t.getTotalPages();
				return n === 0 ? {
					success: !1,
					error: "No pages rendered (headless mode or empty document)."
				} : {
					success: !1,
					error: `Page ${e.pageNumber} does not exist (document has ${n} page${n === 1 ? "" : "s"}).`
				};
			}
			return {
				success: !0,
				data: n.text || "(empty page)"
			};
		}
	},
	{
		name: "read_pages",
		displayName: "Reading pages",
		description: "Read a contiguous range of rendered pages (1-indexed, inclusive). Returns paragraphs across the range, each tagged with paraId. Cheaper than calling read_page repeatedly — single round-trip.",
		inputSchema: {
			type: "object",
			properties: {
				from: {
					type: "number",
					description: "1-indexed start page (inclusive)."
				},
				to: {
					type: "number",
					description: "1-indexed end page (inclusive)."
				}
			},
			required: ["from", "to"]
		},
		handler: (e, t) => {
			let n = t.getPages({
				from: e.from,
				to: e.to
			});
			if (n.length === 0) {
				let n = t.getTotalPages();
				return n === 0 ? {
					success: !1,
					error: "No pages rendered (headless mode or empty document)."
				} : {
					success: !1,
					error: `No pages in range ${e.from}–${e.to} (document has ${n} page${n === 1 ? "" : "s"}).`
				};
			}
			return {
				success: !0,
				data: n.map((e) => `--- Page ${e.pageNumber} ---\n${e.text || "(empty page)"}`).join("\n\n")
			};
		}
	},
	{
		name: "find_text",
		displayName: "Finding text",
		description: "Locate paragraphs containing `query`. Returns up to `limit` handles, each with `paraId`, the matched substring, and surrounding context. Pass any returned `paraId` (and the `match` as `search`) to add_comment / suggest_change.",
		inputSchema: {
			type: "object",
			properties: {
				query: {
					type: "string",
					description: "Text to find (substring match)."
				},
				caseSensitive: {
					type: "boolean",
					description: "Default: false."
				},
				limit: {
					type: "number",
					description: "Max paragraphs to return. Default: 20."
				}
			},
			required: ["query"]
		},
		handler: (e, t) => {
			let n = t.findText(e.query, {
				caseSensitive: e.caseSensitive,
				limit: e.limit
			});
			return n.length === 0 ? {
				success: !0,
				data: "No matches."
			} : {
				success: !0,
				data: n
			};
		}
	},
	{
		name: "read_comments",
		displayName: "Reading comments",
		description: "List all comments in the document with their paragraph anchors.",
		inputSchema: {
			type: "object",
			properties: {}
		},
		handler: (e, t) => {
			let n = t.getComments();
			return n.length === 0 ? {
				success: !0,
				data: "No comments."
			} : {
				success: !0,
				data: n.map((e) => `[Comment #${e.id}] ${e.author}: "${e.text}"` + (e.anchoredText ? ` (anchored to: "${e.anchoredText}")` : "") + (e.replies.length > 0 ? "\n" + e.replies.map((e) => `  Reply by ${e.author}: "${e.text}"`).join("\n") : "")).join("\n")
			};
		}
	},
	{
		name: "read_changes",
		displayName: "Reading changes",
		description: "List tracked changes (insertions / deletions) currently in the document.",
		inputSchema: {
			type: "object",
			properties: {}
		},
		handler: (e, t) => {
			let n = t.getChanges();
			return n.length === 0 ? {
				success: !0,
				data: "No tracked changes."
			} : {
				success: !0,
				data: n.map((e) => `[Change #${e.id}] ${e.type} by ${e.author}: "${e.text}"`).join("\n")
			};
		}
	},
	{
		name: "add_comment",
		displayName: "Adding comment",
		description: "Attach a comment to a paragraph, optionally anchored to a unique phrase within it. The user sees it instantly in the comments sidebar.",
		inputSchema: {
			type: "object",
			properties: {
				paraId: {
					type: "string",
					description: "Paragraph id from read_document / find_text."
				},
				text: {
					type: "string",
					description: "Comment body."
				},
				search: {
					type: "string",
					description: "Optional: anchor to this exact phrase within the paragraph. Must be unique."
				}
			},
			required: ["paraId", "text"]
		},
		handler: (e, t) => {
			let n = t.addComment({
				paraId: e.paraId,
				text: e.text,
				search: e.search
			});
			return n === null ? {
				success: !1,
				error: "Could not add comment. The paraId may not exist, or `search` is missing / ambiguous."
			} : {
				success: !0,
				data: `Comment ${n} added on ${e.paraId}.`
			};
		}
	},
	{
		name: "suggest_change",
		displayName: "Suggesting change",
		description: "Suggest a tracked change. Three modes: (1) replacement — `search` non-empty, `replaceWith` non-empty; (2) deletion — `search` non-empty, `replaceWith` empty; (3) insertion at paragraph end — `search` empty, `replaceWith` non-empty. The user can accept or reject in the editor UI.",
		inputSchema: {
			type: "object",
			properties: {
				paraId: {
					type: "string",
					description: "Paragraph id from read_document / find_text."
				},
				search: {
					type: "string",
					description: "Phrase to find (must be unique). Empty string = insert at paragraph end."
				},
				replaceWith: {
					type: "string",
					description: "Replacement text. Empty string = delete the matched phrase."
				}
			},
			required: [
				"paraId",
				"search",
				"replaceWith"
			]
		},
		handler: (e, t) => t.proposeChange({
			paraId: e.paraId,
			search: e.search,
			replaceWith: e.replaceWith
		}) ? e.search ? e.replaceWith ? {
			success: !0,
			data: `Replacement proposed: "${e.search}" → "${e.replaceWith}" on ${e.paraId}.`
		} : {
			success: !0,
			data: `Deletion proposed: "${e.search}" on ${e.paraId}.`
		} : {
			success: !0,
			data: `Insertion proposed on ${e.paraId}.`
		} : {
			success: !1,
			error: "Could not propose change. Possible causes: paraId not found; search missing or ambiguous; or the target overlaps an existing tracked change."
		}
	},
	Fe,
	Ie,
	{
		name: "insert_break",
		displayName: "Inserting break",
		description: "Insert a break after a paragraph (located by paraId). `type: \"page\"` adds a page break so following content starts on a new page. `type: \"sectionNextPage\"` starts a new document section on a new page. `type: \"sectionContinuous\"` starts a new section on the same page (use this to change column layout or margins from that point onward). The target paragraph becomes the end of its page/section and the next paragraph begins the new one. Direct edit — not a tracked change.",
		inputSchema: {
			type: "object",
			properties: {
				paraId: {
					type: "string",
					description: "Paragraph id from read_document / find_text. The break is inserted after this paragraph."
				},
				type: {
					type: "string",
					enum: [...H],
					description: "Break kind: \"page\" (page break), \"sectionNextPage\" (new section, new page), or \"sectionContinuous\" (new section, same page)."
				}
			},
			required: ["paraId", "type"]
		},
		handler: (e, t) => H.includes(e.type) ? t.insertBreak({
			paraId: e.paraId,
			type: e.type
		}) ? {
			success: !0,
			data: `${Le[e.type]} inserted after ${e.paraId}.`
		} : {
			success: !1,
			error: `Could not insert break. paraId "${e.paraId}" not found.`
		} : {
			success: !1,
			error: `Invalid break type "${e.type}". Must be one of: ${H.join(", ")}.`
		}
	},
	{
		name: "reply_comment",
		displayName: "Replying to comment",
		description: "Reply to an existing comment by id. Threaded under the original.",
		inputSchema: {
			type: "object",
			properties: {
				commentId: {
					type: "number",
					description: "Comment id from read_comments."
				},
				text: {
					type: "string",
					description: "Reply body."
				}
			},
			required: ["commentId", "text"]
		},
		handler: (e, t) => {
			let n = t.replyTo(e.commentId, { text: e.text });
			return n === null ? {
				success: !1,
				error: `Comment #${e.commentId} not found.`
			} : {
				success: !0,
				data: `Reply ${n} added to comment ${e.commentId}.`
			};
		}
	},
	{
		name: "resolve_comment",
		displayName: "Resolving comment",
		description: "Mark a comment as resolved (done).",
		inputSchema: {
			type: "object",
			properties: { commentId: {
				type: "number",
				description: "Comment id from read_comments."
			} },
			required: ["commentId"]
		},
		handler: (e, t) => (t.resolveComment(e.commentId), {
			success: !0,
			data: `Comment ${e.commentId} resolved.`
		})
	},
	{
		name: "scroll",
		displayName: "Scrolling",
		description: "Scroll the editor to a paragraph by paraId. Does not move the user's cursor.",
		inputSchema: {
			type: "object",
			properties: { paraId: {
				type: "string",
				description: "Paragraph id from read_document / find_text."
			} },
			required: ["paraId"]
		},
		handler: (e, t) => t.scrollTo(e.paraId) ? {
			success: !0,
			data: `Scrolled to ${e.paraId}.`
		} : {
			success: !1,
			error: `paraId ${e.paraId} not found.`
		}
	}
];
function Re(e, t, n) {
	let r = U.find((t) => t.name === e);
	if (!r) return {
		success: !1,
		error: `Unknown tool: ${e}`
	};
	try {
		return r.handler(t, n);
	} catch (e) {
		return {
			success: !1,
			error: e instanceof Error ? e.message : String(e)
		};
	}
}
function ze(e) {
	let t = U.find((t) => t.name === e);
	if (t?.displayName) return t.displayName;
	let n = e.replace(/_/g, " ");
	return n.charAt(0).toUpperCase() + n.slice(1);
}
function Be() {
	return U.map((e) => ({
		type: "function",
		function: {
			name: e.name,
			description: e.description,
			parameters: e.inputSchema
		}
	}));
}
//#endregion
//#region src/utils.ts
function W(e) {
	return e.type === "insertion" || e.type === "deletion" || e.type === "moveFrom" || e.type === "moveTo";
}
function G(e) {
	let t = [];
	for (let n of e) n.type === "run" ? t.push(E(n)) : n.type === "hyperlink" && t.push(T(n));
	return t.join("");
}
function K(e, t, n) {
	let r = 0;
	for (let i of e) if (i.type === "paragraph") {
		if (n(i, r) === !1) return;
		r++;
	} else if (i.type === "table") {
		for (let e of i.rows) for (let t of e.cells) for (let e of t.content) if (e.type === "paragraph") {
			if (n(e, r) === !1) return;
			r++;
		}
	} else t && r++;
}
function q(e, t) {
	K(e.content, !0, t);
}
function J(e, t) {
	K(e.content, !1, t);
}
//#endregion
//#region src/content.ts
function Y(e, t = {}) {
	let { fromIndex: n, toIndex: r, includeTrackedChanges: i = !0, includeCommentAnchors: a = !0 } = t, o = [], s = 0;
	for (let t of e.content) if (t.type === "paragraph") X(s, n, r) && o.push(Ve(t, s, i, a)), s++;
	else if (t.type === "table") {
		X(s, n, r) && o.push(He(t, s, i, a));
		for (let e of t.rows) for (let t of e.cells) for (let e of t.content) e.type === "paragraph" && s++;
	} else s++;
	return o;
}
function X(e, t, n) {
	return (t === void 0 || e >= t) && (n === void 0 || e <= n);
}
function Ve(e, t, n, r) {
	let i = Z(e, n, r), a = e.formatting?.styleId, o = e.paraId;
	return D(a) ? {
		type: "heading",
		index: t,
		paraId: o,
		level: O(a) ?? 1,
		text: i
	} : e.listRendering ? {
		type: "list-item",
		index: t,
		paraId: o,
		text: i,
		listLevel: e.listRendering.level ?? 0,
		listType: e.listRendering.isBullet ? "bullet" : "number"
	} : {
		type: "paragraph",
		index: t,
		paraId: o,
		text: i
	};
}
function He(e, t, n, r) {
	let i = [], a = [];
	for (let t of e.rows) {
		let e = [], o = [];
		for (let i of t.cells) {
			let t = [], a;
			for (let e of i.content) e.type === "paragraph" && (a === void 0 && (a = e.paraId), t.push(Z(e, n, r)));
			e.push(t.join("\n")), o.push(a);
		}
		i.push(e), a.push(o);
	}
	return {
		type: "table",
		index: t,
		rows: i,
		cellParaIds: a
	};
}
function Ue(e) {
	let t = [], n = (e) => e.paraId ? e.paraId : String(e.index);
	for (let r of e) switch (r.type) {
		case "heading":
			t.push(`[${n(r)}] (h${r.level}) ${r.text}`);
			break;
		case "paragraph":
			t.push(`[${n(r)}] ${r.text}`);
			break;
		case "list-item": {
			let e = "  ".repeat(r.listLevel), i = r.listType === "bullet" ? "•" : "-";
			t.push(`[${n(r)}] ${e}${i} ${r.text}`);
			break;
		}
		case "table": {
			let e = r.index;
			for (let n = 0; n < r.rows.length; n++) for (let i = 0; i < r.rows[n].length; i++) {
				let a = r.rows[n][i].split("\n"), o = r.cellParaIds?.[n]?.[i];
				for (let r = 0; r < a.length; r++) {
					let s = r === 0 && o ? o : String(e);
					t.push(`[${s}] (table, row ${n + 1}, col ${i + 1}) ${a[r]}`), e++;
				}
			}
			break;
		}
	}
	return t.join("\n");
}
function Z(e, t, n) {
	let r = [], i = /* @__PURE__ */ new Set();
	for (let a of e.content) {
		if (a.type === "commentRangeStart" && n) {
			i.add(a.id), r.push(`[comment:${a.id}]`);
			continue;
		}
		if (a.type === "commentRangeEnd" && n) {
			i.has(a.id) && (i.delete(a.id), r.push("[/comment]"));
			continue;
		}
		if (a.type === "run") r.push(E(a));
		else if (a.type === "hyperlink") r.push(T(a));
		else if (W(a)) {
			let e = G(a.content);
			a.type === "insertion" || a.type === "moveTo" ? t && r.push(`[+${e}+]{by:${a.info.author}}`) : t ? r.push(`[-${e}-]{by:${a.info.author}}`) : r.push(e);
		}
	}
	return r.join("");
}
//#endregion
//#region src/textSearch.ts
function We(e) {
	let t = [], n = 0;
	for (let r = 0; r < e.content.length; r++) {
		let i = e.content[r];
		if (i.type === "run") {
			let e = E(i);
			t.push({
				contentIndex: r,
				run: i,
				text: e,
				startPos: n
			}), n += e.length;
		} else if (i.type === "hyperlink") for (let e = 0; e < i.children.length; e++) {
			let a = i.children[e];
			if (a.type === "run") {
				let e = E(a);
				t.push({
					contentIndex: r,
					run: a,
					text: e,
					startPos: n
				}), n += e.length;
			}
		}
		else if (W(i)) {
			if (i.type === "insertion" || i.type === "moveTo") continue;
			for (let e = 0; e < i.content.length; e++) {
				let a = i.content[e];
				if (a.type === "run") {
					let e = E(a);
					t.push({
						contentIndex: r,
						run: a,
						text: e,
						startPos: n
					}), n += e.length;
				} else if (a.type === "hyperlink") {
					for (let e of a.children) if (e.type === "run") {
						let i = E(e);
						t.push({
							contentIndex: r,
							run: e,
							text: i,
							startPos: n
						}), n += i.length;
					}
				}
			}
		}
	}
	return t;
}
function Ge(e) {
	return We(e).map((e) => e.text).join("");
}
//#endregion
//#region src/discovery.ts
function Ke(e, t, n) {
	let r = /* @__PURE__ */ new Map(), i = (e, t, n) => {
		let i = null;
		for (let a of e.content) if (W(a)) {
			i === null && (i = Ge(e));
			let o = G(a.content), s = a.info.id, c = `${n.noteType ?? "body"}:${n.noteId ?? ""}:${t}:${s}`, l = r.get(c);
			l ? l.text += o : r.set(c, {
				id: s,
				type: a.type,
				author: a.info.author,
				date: a.info.date ?? null,
				text: o,
				context: i,
				paragraphIndex: t,
				...n.noteId === void 0 ? {} : { noteId: n.noteId },
				...n.noteType === void 0 ? {} : { noteType: n.noteType }
			});
		}
	};
	if (q(e, (e, t) => i(e, t, {})), t?.includeFootnotes && n?.footnotes) for (let e of n.footnotes) J(e, (t, n) => i(t, n, {
		noteId: e.id,
		noteType: "footnote"
	}));
	if (t?.includeEndnotes && n?.endnotes) for (let e of n.endnotes) J(e, (t, n) => i(t, n, {
		noteId: e.id,
		noteType: "endnote"
	}));
	return Array.from(r.values()).filter((e) => !(t?.author && e.author !== t.author || t?.type && e.type !== t.type));
}
function qe(e, t) {
	let n = e.comments ?? [];
	if (n.length === 0) return [];
	let r = Ye(e), i = [], a = /* @__PURE__ */ new Map();
	for (let e of n) if (e.parentId !== void 0) {
		let t = a.get(e.parentId) ?? [];
		t.push(e), a.set(e.parentId, t);
	} else i.push(e);
	return i.map((e) => {
		let t = r.get(e.id), n = (a.get(e.id) ?? []).map((e) => ({
			id: e.id,
			author: e.author,
			date: e.date ?? null,
			text: Je(e)
		}));
		return {
			id: e.id,
			author: e.author,
			date: e.date ?? null,
			text: Je(e),
			anchoredText: t?.text ?? "",
			paragraphIndex: t?.paragraphIndex ?? -1,
			replies: n,
			done: e.done ?? !1
		};
	}).filter((e) => !(t?.author && e.author !== t.author || t?.done !== void 0 && e.done !== t.done));
}
function Je(e) {
	return e.content.map((e) => Ge(e)).join("\n");
}
function Ye(e) {
	let t = /* @__PURE__ */ new Map(), n = /* @__PURE__ */ new Map();
	return q(e, (e, r) => {
		for (let i of e.content) if (i.type === "commentRangeStart") n.set(i.id, {
			paragraphIndex: r,
			parts: []
		});
		else if (i.type === "commentRangeEnd") {
			let e = n.get(i.id);
			e && (t.set(i.id, {
				text: e.parts.join(""),
				paragraphIndex: e.paragraphIndex
			}), n.delete(i.id));
		} else if (i.type === "run") {
			let e = E(i);
			for (let t of n.values()) t.parts.push(e);
		} else if (i.type === "hyperlink") {
			let e = i.children.filter((e) => e.type === "run").map(E).join("");
			for (let t of n.values()) t.parts.push(e);
		} else if (W(i)) {
			if (i.type === "insertion" || i.type === "moveTo") continue;
			let e = G(i.content);
			for (let t of n.values()) t.parts.push(e);
		}
	}), t;
}
//#endregion
//#region src/bridge.ts
function Q(e) {
	return !e || e.length === 0 ? "" : e.map((e) => {
		let t = e;
		return t?.content ? t.content.map((e) => e.content?.map((e) => e.text || "").join("") || "").join("") : "";
	}).join("\n");
}
function $(e) {
	let t = e.getEditorRef();
	if (t) {
		let e = t.getDocument();
		if (e?.package?.document) return e.package.document;
	}
	return e.getDocument()?.package?.document ?? null;
}
function Xe(e, t = "AI") {
	function n(e) {
		return e ?? t;
	}
	return {
		getContentAsText(t) {
			let n = $(e);
			return n ? Ue(Y(n, t)) : "";
		},
		getContent(t) {
			let n = $(e);
			return n ? Y(n, t) : [];
		},
		getComments(t) {
			let n = $(e);
			if (!n) return [];
			let r = qe(n, t);
			if (r.length > 0) return r;
			let i = e.getComments();
			if (i.length === 0) return [];
			let a = /* @__PURE__ */ new Map(), o = [];
			for (let e of i) if (e.parentId) {
				let t = a.get(e.parentId);
				t ? t.push(e) : a.set(e.parentId, [e]);
			} else o.push(e);
			let s = [];
			for (let e of o) {
				if (t?.author && e.author !== t.author || t?.done !== void 0 && (e.done ?? !1) !== t.done) continue;
				let n = a.get(e.id) ?? [];
				s.push({
					id: e.id,
					author: e.author,
					date: e.date ?? null,
					text: Q(e.content),
					anchoredText: "",
					paragraphIndex: -1,
					replies: n.map((e) => ({
						id: e.id,
						author: e.author,
						date: e.date ?? null,
						text: Q(e.content)
					})),
					done: e.done ?? !1
				});
			}
			return s;
		},
		getChanges(t) {
			let n = $(e);
			return n ? Ke(n, t) : [];
		},
		findText(t, n) {
			return e.findInDocument(t, n);
		},
		getSelection() {
			return e.getSelectionInfo();
		},
		addComment(t) {
			return e.addComment({
				paraId: t.paraId,
				text: t.text,
				author: n(t.author),
				search: t.search
			});
		},
		replyTo(t, r) {
			return e.replyToComment(t, r.text, n(r.author));
		},
		resolveComment(t) {
			e.resolveComment(t);
		},
		proposeChange(t) {
			return e.proposeChange({
				paraId: t.paraId,
				search: t.search,
				replaceWith: t.replaceWith,
				author: n(t.author)
			});
		},
		applyFormatting(t) {
			return e.applyFormatting({
				paraId: t.paraId,
				search: t.search,
				marks: t.marks
			});
		},
		setParagraphStyle(t) {
			return e.setParagraphStyle({
				paraId: t.paraId,
				styleId: t.styleId
			});
		},
		insertBreak(t) {
			return e.insertBreak({
				paraId: t.paraId,
				type: t.type
			});
		},
		getPage(t) {
			return e.getPageContent(t);
		},
		getPages(t) {
			let n = e.getTotalPages(), r = Math.max(1, Math.min(t.from, n)), i = Math.max(r, Math.min(t.to, n)), a = [];
			for (let t = r; t <= i; t++) {
				let n = e.getPageContent(t);
				n && a.push(n);
			}
			return a;
		},
		getTotalPages() {
			return e.getTotalPages();
		},
		getCurrentPage() {
			return e.getCurrentPage();
		},
		scrollTo(t, n) {
			return e.scrollToParaId(t, n);
		},
		onContentChange(t) {
			return e.onContentChange(() => {
				let n = $(e), r = n ? qe(n) : [], i = n ? Ke(n) : [];
				try {
					t({
						commentCount: r.length,
						changeCount: i.length,
						comments: r,
						changes: i
					});
				} catch (e) {
					console.error("onContentChange listener threw:", e);
				}
			});
		},
		onSelectionChange(t) {
			return e.onSelectionChange(() => {
				try {
					t(e.getSelectionInfo());
				} catch (e) {
					console.error("onSelectionChange listener threw:", e);
				}
			});
		}
	};
}
//#endregion
//#region src/vue/composables/useAgentBridge.ts
var Ze = Be();
function Qe(e) {
	let { editorRef: t, author: r = "AI" } = e, i = n(() => t.value ? Xe(t.value, y(r)) : null);
	function a(e, t) {
		let n = i.value;
		return n ? Re(e, t, n) : {
			success: !1,
			error: "Editor not ready"
		};
	}
	return {
		executeToolCall: a,
		toolSchemas: Ze
	};
}
//#endregion
export { ye as AIContextMenu, Pe as AIResponsePreview, ae as AgentChatLog, ce as AgentComposer, M as AgentPanel, ue as AgentSuggestionChip, L as AgentTimeline, ze as getToolDisplayName, Qe as useAgentBridge };
