export{a as RenderedDomContextImpl,b as createRenderedDomContext}from'./chunk-CL5ZW2RS.mjs';import {a}from'./chunk-BMBP5UFA.mjs';import ie,{forwardRef,useState,useRef,useMemo,useSyncExternalStore,useEffect,useCallback,useImperativeHandle,cloneElement}from'react';import {TextSelection}from'prosemirror-state';import {PluginLifecycleManager,injectStyles}from'@npde/docx-editor-core';import {jsx,jsxs}from'react/jsx-runtime';import*as ke from'@npde/docx-editor-core/prosemirror/template/prosemirror-plugin';var z={position:"right",defaultSize:280,minSize:200,maxSize:500,resizable:true,collapsible:true,defaultCollapsed:false},j=injectStyles,J=`
.plugin-host {
  display: flex;
  width: 100%;
  height: 100%;
  overflow: visible;
  position: relative;
}

.plugin-host-editor {
  flex: 1;
  display: flex;
  flex-direction: column;
  min-width: 0;
  overflow: visible;
}


.plugin-panels-left,
.plugin-panels-right {
  display: flex;
  flex-direction: column;
  flex-shrink: 0;
  background: #f8f9fa;
  border-color: #e9ecef;
}

.plugin-panels-left {
  border-right: 1px solid #e9ecef;
}

.plugin-panels-right {
  border-left: 1px solid #e9ecef;
}

.plugin-panels-bottom {
  border-top: 1px solid #e9ecef;
  background: #f8f9fa;
}

.plugin-panel {
  position: relative;
  display: flex;
  flex-direction: column;
  overflow: hidden;
  transition: width 0.2s ease, height 0.2s ease;
}

.plugin-panel.collapsed {
  overflow: visible;
}

.plugin-panel-toggle {
  display: flex;
  align-items: center;
  gap: 4px;
  padding: 6px 8px;
  background: transparent;
  border: none;
  cursor: pointer;
  font-size: 12px;
  color: #6c757d;
  white-space: nowrap;
}

.plugin-panel.collapsed .plugin-panel-toggle {
  writing-mode: vertical-rl;
  text-orientation: mixed;
  flex-direction: column;
  height: 100%;
  padding: 8px 6px;
}

.plugin-panel-toggle:hover {
  background: #e9ecef;
  color: #495057;
}

.plugin-panel-toggle-icon {
  font-weight: bold;
  font-size: 14px;
}

.plugin-panel.collapsed .plugin-panel-toggle-icon {
  transform: rotate(90deg);
}

.plugin-panel-toggle-label {
  font-weight: 500;
}

.plugin-panel-content {
  flex: 1;
  overflow: auto;
}

/* Right panel rendered inside viewport - scrolls with content */
.plugin-panel-in-viewport {
  position: absolute;
  top: 0;
  /* Position is set dynamically via inline styles based on page edge */
  width: 220px;
  pointer-events: auto;
  z-index: 10;
  overflow: visible;
}

.plugin-panel-in-viewport.collapsed {
  width: 32px;
}

.plugin-panel-in-viewport .plugin-panel-toggle {
  position: sticky;
  top: 0;
  background: rgba(255, 255, 255, 0.95);
  border-radius: 4px;
  box-shadow: 0 1px 3px rgba(0, 0, 0, 0.1);
}

.plugin-panel-in-viewport-content {
  overflow: visible;
  position: relative;
}

/* Plugin overlay container for rendering highlights/decorations */
.plugin-overlays-container {
  position: absolute;
  top: 0;
  left: 0;
  right: 0;
  bottom: 0;
  pointer-events: none;
  overflow: visible;
  z-index: 5;
}

.plugin-overlay {
  position: absolute;
  top: 0;
  left: 0;
  right: 0;
  bottom: 0;
  pointer-events: none;
}

/* Individual overlay children manage their own pointer-events.
   Do NOT set pointer-events: auto here \u2014 it overrides overlay containers
   that need pointer-events: none to let clicks pass through to the editor. */
`,ve=forwardRef(function({plugins:d,children:g,className:p=""},u){let[n,y]=useState(null),f=useRef(g.props);f.current=g.props;let[a,S]=useState(null),i=useMemo(()=>new PluginLifecycleManager,[]),m=useSyncExternalStore(i.subscribe,i.getSnapshot),[h,T]=useState(()=>{let t=new Set;for(let e of d)({...z,...e.panelConfig}).defaultCollapsed&&t.add(e.id);return t}),[v]=useState(()=>{let t=new Map;for(let e of d){let o={...z,...e.panelConfig};t.set(e.id,o.defaultSize);}return t});useEffect(()=>{if(!n)return;let t=d.map(e=>({id:e.id,styles:e.styles,initialize:e.initialize,onStateChange:e.onStateChange,destroy:e.destroy}));return i.initialize(t,n),()=>{i.destroy();}},[i,n,d]),useEffect(()=>{let t=d.filter(e=>e.styles).map(e=>j(e.id,e.styles));return ()=>t.forEach(e=>e())},[d]),useEffect(()=>{if(!n?.dom)return;let t=()=>{i.updateStates(n);},e=null,o=()=>{e&&cancelAnimationFrame(e),e=requestAnimationFrame(t);};t();let s=n.dom;s.addEventListener("input",o),s.addEventListener("focus",t),s.addEventListener("click",t);let c=n.dispatch.bind(n);return n.dispatch=x=>{c(x),o();},()=>{s.removeEventListener("input",o),s.removeEventListener("focus",t),s.removeEventListener("click",t),e&&cancelAnimationFrame(e),n.dispatch=c;}},[n,i]),useEffect(()=>j("plugin-host-base",J),[]);let O=useCallback(t=>{if(!n)return;if(n.coordsAtPos(t)){n.dom.scrollIntoView({block:"center",inline:"nearest"});let{state:o}=n,s=o.doc.resolve(Math.min(t,o.doc.content.size)),c=o.tr.setSelection(TextSelection.near(s));n.dispatch(c),n.focus();}},[n]),k=useCallback((t,e)=>{if(!n)return;let{state:o}=n,s=o.doc.content.size,c=Math.max(0,Math.min(t,s)),x=Math.max(0,Math.min(e,s)),R=o.tr.setSelection(TextSelection.create(o.doc,c,x));n.dispatch(R),n.focus();},[n]),_=useCallback(t=>i.getPluginState(t),[i]),A=useCallback((t,e)=>{i.setPluginState(t,e);},[i]),F=useCallback(()=>{n&&i.updateStates(n);},[n,i]);useImperativeHandle(u,()=>({getPluginState:_,setPluginState:A,getEditorView:()=>n,refreshPluginStates:F}),[_,A,n,F]);let Y=useMemo(()=>{let t=[];for(let e of d)e.proseMirrorPlugins&&t.push(...e.proseMirrorPlugins);return t},[d]),M=useCallback(t=>{T(e=>{let o=new Set(e);return o.has(t)?o.delete(t):o.add(t),o});},[]),[$,N]=useState(null);useEffect(()=>{if(!a){N(null);return}let t=()=>{let s=a.pagesContainer,c=s.querySelector(".layout-page");if(!c){N(null);return}let x=a.getContainerOffset(),R=c.getBoundingClientRect(),D=s.getBoundingClientRect(),V=(R.right-D.left)/a.zoom,le=x.x+V+5;N(le);};t();let e=()=>{requestAnimationFrame(t);};window.addEventListener("resize",e);let o=new ResizeObserver(()=>{requestAnimationFrame(t);});return o.observe(a.pagesContainer),()=>{window.removeEventListener("resize",e),o.disconnect();}},[a]);let G=useMemo(()=>{let t=[];if(a){for(let e of d)if(e.renderOverlay){let o=m.states.get(e.id);t.push(jsx("div",{className:"plugin-overlay","data-plugin-id":e.id,children:e.renderOverlay(a,o,n)},`overlay-${e.id}`));}}for(let e of d){if(!e.Panel||(e.panelConfig?.position??"right")!=="right")continue;let s={...z,...e.panelConfig},c=h.has(e.id),x=v.get(e.id)??s.defaultSize,R=e.Panel,D=m.states.get(e.id),V=$!==null?`${$}px`:"calc(50% + 428px)";t.push(jsxs("div",{className:`plugin-panel-in-viewport ${c?"collapsed":""}`,style:{width:c?"32px":`${x}px`,left:V},"data-plugin-id":e.id,children:[s.collapsible&&jsx("button",{className:"plugin-panel-toggle",onClick:()=>M(e.id),title:c?`Show ${e.name}`:`Hide ${e.name}`,"aria-label":c?`Show ${e.name}`:`Hide ${e.name}`,children:jsx("span",{className:"plugin-panel-toggle-icon",children:c?"\u2039":"\u203A"})}),!c&&a&&jsx("div",{className:"plugin-panel-in-viewport-content",children:jsx(R,{editorView:n,doc:n?.state.doc??null,scrollToPosition:O,selectRange:k,pluginState:D,panelWidth:x,renderedDomContext:a})})]},`panel-overlay-${e.id}`));}return t.length>0?t:null},[a,d,m.version,n,h,v,O,k,M,$]),q=useMemo(()=>{let t=[];for(let e of d){if(!e.getSidebarItems)continue;let o=m.states.get(e.id),s={editorView:n,renderedDomContext:a,anchorPositions:new Map,zoom:a?.zoom??1},c=e.getSidebarItems(o,s);t.push(...c);}return t},[d,m.version,n,a]),U=useCallback(t=>{S(t);let e=f.current?.onRenderedDomContextReady;typeof e=="function"&&e(t);},[]),se=useMemo(()=>cloneElement(g,{externalPlugins:Y,pluginOverlays:G,pluginSidebarItems:q,pluginRenderedDomContext:a,onRenderedDomContextReady:U,onEditorViewReady:t=>{y(t);let e=f.current?.onEditorViewReady;typeof e=="function"&&e(t);}}),[g,Y,G,q,a,U]),H=useMemo(()=>{let t=[],e=[],o=[];for(let s of d){if(!s.Panel)continue;let c=s.panelConfig?.position??"right";c==="left"?t.push(s):c==="bottom"?o.push(s):e.push(s);}return {left:t,right:e,bottom:o}},[d]),W=t=>{if(!t.Panel)return null;let e={...z,...t.panelConfig},o=h.has(t.id),s=v.get(t.id)??e.defaultSize,c=t.Panel,x=m.states.get(t.id);return jsxs("div",{className:`plugin-panel plugin-panel-${e.position} ${o?"collapsed":""}`,style:{[e.position==="bottom"?"height":"width"]:o?"32px":`${s}px`,minWidth:e.position!=="bottom"?o?"32px":`${e.minSize}px`:void 0,maxWidth:e.position!=="bottom"?`${e.maxSize}px`:void 0,minHeight:e.position==="bottom"?o?"32px":`${e.minSize}px`:void 0,maxHeight:e.position==="bottom"?`${e.maxSize}px`:void 0},"data-plugin-id":t.id,children:[e.collapsible&&jsxs("button",{className:"plugin-panel-toggle",onClick:()=>M(t.id),title:o?`Show ${t.name}`:`Hide ${t.name}`,"aria-label":o?`Show ${t.name}`:`Hide ${t.name}`,children:[jsx("span",{className:"plugin-panel-toggle-icon",children:o?"\u203A":"\u2039"}),o&&jsx("span",{className:"plugin-panel-toggle-label",children:t.name})]}),!o&&jsx("div",{className:"plugin-panel-content",children:jsx(c,{editorView:n,doc:n?.state.doc??null,scrollToPosition:O,selectRange:k,pluginState:x,panelWidth:s,renderedDomContext:a??null})})]},t.id)};return jsxs("div",{className:`plugin-host ${p}`,children:[H.left.length>0&&jsx("div",{className:"plugin-panels-left",children:H.left.map(W)}),jsxs("div",{className:"plugin-host-editor",children:[se,H.bottom.length>0&&jsx("div",{className:"plugin-panels-bottom",children:H.bottom.map(W)})]})]})});var l={};a(l,ke);var Se={variable:"rgba(245, 158, 11, 0.3)",sectionStart:"rgba(59, 130, 246, 0.3)",sectionEnd:"rgba(59, 130, 246, 0.3)",invertedStart:"rgba(139, 92, 246, 0.3)",raw:"rgba(239, 68, 68, 0.3)"},Te={variable:"rgba(245, 158, 11, 0.5)",sectionStart:"rgba(59, 130, 246, 0.5)",sectionEnd:"rgba(59, 130, 246, 0.5)",invertedStart:"rgba(139, 92, 246, 0.5)",raw:"rgba(239, 68, 68, 0.5)"};function Z({context:r,tags:d,hoveredId:g,selectedId:p,onHover:u,onSelect:n}){let[y,f]=useState(0),a=useCallback(()=>{let i=r.getContainerOffset(),m=[];for(let h of d){let T=r.getRectsForRange(h.from,h.to);for(let v of T)m.push({tagId:h.id,tagType:h.type,x:v.x+i.x,y:v.y+i.y,width:v.width,height:v.height});}return m},[r,d]),S=useMemo(()=>a(),[a,y]);return useEffect(()=>{let i=()=>{requestAnimationFrame(()=>f(m=>m+1));};return window.addEventListener("resize",i),()=>window.removeEventListener("resize",i)},[]),useEffect(()=>{let i=new ResizeObserver(()=>{requestAnimationFrame(()=>f(m=>m+1));});return i.observe(r.pagesContainer),()=>i.disconnect()},[r.pagesContainer]),S.length===0?null:jsx("div",{className:"template-highlight-overlay",children:S.map((i,m)=>{let h=i.tagId===g,T=i.tagId===p,v=h||T?Te[i.tagType]:Se[i.tagType];return jsx("div",{className:`template-highlight ${h?"hovered":""} ${T?"selected":""}`,style:{position:"absolute",left:i.x,top:i.y,width:i.width,height:i.height,backgroundColor:v,borderRadius:2,cursor:"pointer"},onMouseEnter:()=>u?.(i.tagId),onMouseLeave:()=>u?.(void 0),onClick:()=>n?.(i.tagId)},`${i.tagId}-${m}`)})})}var ee=`
.template-highlight-overlay {
  position: absolute;
  top: 0;
  left: 0;
  right: 0;
  bottom: 0;
  pointer-events: none;
  overflow: visible;
}

.template-highlight {
  pointer-events: auto;
  transition: background-color 0.1s ease;
}

.template-highlight:hover,
.template-highlight.hovered {
  filter: brightness(0.9);
}

.template-highlight.selected {
  box-shadow: 0 0 0 2px rgba(59, 130, 246, 0.6);
}
`;var we={variable:"#f59e0b",sectionStart:"#3b82f6",sectionEnd:"#3b82f6",invertedStart:"#8b5cf6",raw:"#ef4444"};function Ee(r){switch(r){case "sectionStart":return "LOOP / IF";case "invertedStart":return "IF NOT";case "raw":return "HTML";default:return ""}}var ne=`
.template-annotation-chip {
  display: inline-flex;
  flex-wrap: wrap;
  align-items: center;
  gap: 4px;
  padding: 5px 10px;
  background: white;
  border: 1px solid #e2e8f0;
  border-left: 3px solid #6c757d;
  border-radius: 4px;
  font-size: 11px;
  cursor: pointer;
  box-shadow: 0 1px 3px rgba(0, 0, 0, 0.06);
  max-width: 200px;
}

.template-annotation-chip:hover,
.template-annotation-chip.hovered {
  box-shadow: 0 2px 6px rgba(0, 0, 0, 0.1);
  border-color: #cbd5e1;
}

.template-annotation-chip.selected {
  box-shadow: 0 0 0 2px rgba(59, 130, 246, 0.5);
}

.template-chip-badge {
  font-size: 9px;
  font-weight: 600;
  padding: 1px 5px;
  border-radius: 3px;
  color: white;
  text-transform: uppercase;
  letter-spacing: 0.3px;
}

.template-chip-dot {
  font-size: 8px;
}

.template-chip-name {
  color: #334155;
  font-weight: 500;
}

.template-chip-nested {
  display: flex;
  flex-wrap: wrap;
  gap: 4px;
  width: 100%;
  margin-top: 4px;
  padding-top: 4px;
  border-top: 1px solid rgba(0, 0, 0, 0.06);
}

.template-nested-var {
  font-size: 10px;
  color: #64748b;
  background: rgba(0, 0, 0, 0.04);
  padding: 2px 6px;
  border-radius: 3px;
}

.template-nested-var:hover {
  background: rgba(59, 130, 246, 0.15);
  color: #1e40af;
}
`;function oe({tag:r,isHovered:d,measureRef:g,onHover:p,onSelect:u}){let n=Ee(r.type),y=we[r.type],f=r.type==="sectionStart"||r.type==="invertedStart";return jsxs("div",{ref:g,style:{display:"flex",alignItems:"flex-start"},children:[jsx("div",{style:{width:20,height:1,background:d?"#3b82f6":"#d0d0d0",marginTop:12,marginRight:4,flexShrink:0}}),jsxs("div",{className:`template-annotation-chip ${d?"hovered":""}`,style:{borderLeftColor:y},onMouseEnter:()=>p(r.id),onMouseLeave:()=>p(void 0),onClick:a=>{a.stopPropagation(),u(r.id);},onMouseDown:a=>a.stopPropagation(),title:f?`${r.rawTag}
Iterates over ${r.name}[]. Access nested properties via ${r.name}.property`:r.rawTag,children:[n&&jsx("span",{className:"template-chip-badge",style:{background:y},children:n}),!n&&jsx("span",{className:"template-chip-dot",style:{color:y},children:"\u25CF"}),jsx("span",{className:"template-chip-name",children:r.name}),f&&r.nestedVars&&r.nestedVars.length>0&&jsx("div",{className:"template-chip-nested",children:r.nestedVars.map((a,S)=>jsx("span",{className:"template-nested-var",title:`Access: ${r.name}.${a}`,children:a.includes(".")?a.split(".").pop():a},S))})]})]})}function re(r,d,g){if(!r)return;(0, l.setSelectedElement)(r,g);let p=d.find(u=>u.id===g);if(p){let u=r.state.tr.setSelection(TextSelection.near(r.state.doc.resolve(p.from)));r.dispatch(u),r.focus();}}function ae(r={}){return {id:"template",name:"Template",proseMirrorPlugins:[(0, l.createTemplatePlugin)()],onStateChange:g=>{let p=l.templatePluginKey.getState(g.state);if(p)return {tags:p.tags,hoveredId:p.hoveredId,selectedId:p.selectedId}},initialize:g=>({tags:[]}),getSidebarItems:(g,p)=>!g||g.tags.length===0?[]:g.tags.filter(n=>n.type!=="sectionEnd"&&!n.insideSection).map(n=>({id:`template-${n.id}`,anchorPos:n.from,priority:10,estimatedHeight:32,render:y=>ie.createElement(oe,{...y,tag:n,isHovered:n.id===g.hoveredId,onHover:f=>{p.editorView&&(0, l.setHoveredElement)(p.editorView,f);},onSelect:f=>re(p.editorView,g.tags,f)})})),renderOverlay:(g,p,u)=>!p||p.tags.length===0?null:ie.createElement(Z,{context:g,tags:p.tags,hoveredId:p.hoveredId,selectedId:p.selectedId,onHover:n=>{u&&(0, l.setHoveredElement)(u,n);},onSelect:n=>re(u,p.tags,n)}),styles:`
${l.TEMPLATE_DECORATION_STYLES}
${ne}
${ee}
`}}var Ce=ae();var export_TEMPLATE_DECORATION_STYLES=l.TEMPLATE_DECORATION_STYLES;var export_createTemplateProseMirrorPlugin=l.createTemplatePlugin;var export_getTemplatePluginTags=l.getTemplateTags;var export_setHoveredElement=l.setHoveredElement;var export_setSelectedElement=l.setSelectedElement;var export_templatePluginKey=l.templatePluginKey;export{J as PLUGIN_HOST_STYLES,ve as PluginHost,export_TEMPLATE_DECORATION_STYLES as TEMPLATE_DECORATION_STYLES,ae as createTemplatePlugin,export_createTemplateProseMirrorPlugin as createTemplateProseMirrorPlugin,export_getTemplatePluginTags as getTemplatePluginTags,export_setHoveredElement as setHoveredElement,export_setSelectedElement as setSelectedElement,Ce as templatePlugin,export_templatePluginKey as templatePluginKey};