#pragma once
#include "dispobject.h"

class CSciXML : public CDispObject
{
public:
	CSciXML(void);
	virtual ~CSciXML(void);
	BOOL			loadFromString(
						LPCTSTR			szString);
	BOOL			initNew();
	BOOL			saveToString(
						LPTSTR		*	szString);
	BOOL			GetrootName(
						LPTSTR			szrootName,
						UINT			nBufferSize);
	void			SetrootName(
						LPCTSTR			szrootName);
	BOOL			GetrootNode(
						IDispatch	**	ppdisproot);
	long			determineChildNodes(
						IDispatch	*	pdispnode);
	BOOL			getChildNode(
						IDispatch	*	pdispnode,
						long			index,
						IDispatch	**	ppdispChild);
	BOOL			getNodeName(
						IDispatch	*	pdispnode,
						LPTSTR			szNodeName,
						UINT			nBufferSize);
	BOOL			getNodeValue(
						IDispatch	*	pdispnode,
						LPTSTR			szNodeValue,
						UINT			nBufferSize);
	BOOL			getNodeValue(
						IDispatch	*	pdispnode,
						LPTSTR		*	szNodeValue);
	// get the number of grating scans
	long			GetNGratingScans(
						long		*	nGratingScans,
						LPTSTR		*	arrGratingScans);
	BOOL			GetChildNodeName(
						IDispatch	*	pdispnode,
						long			index,
						LPTSTR			szChild,
						UINT			nBufferSize);
	// find named child
	BOOL			findNamedChild(
						IDispatch	*	pdispnode,
						LPCTSTR			szChild,
						IDispatch	**	ppdispchild);
	// create a new node
	BOOL			createNode(
						LPCTSTR			sznodeName,
						IDispatch	**	ppdispNode);
	// set node value
	void			setNodeValue(
						IDispatch	*	pdispNode,
						LPCTSTR			szValue);
	// append a child node
	void			appendChildNode(
						IDispatch	*	pdispNode,
						IDispatch	*	pdispChild);
};
