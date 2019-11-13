#pragma once

#include "stdafx.h"

// Global taskpane
static class Globals
{
	static CustomTaskPane* m_pTaskPane;

	void SetTaskpane(CustomTaskPane* pTaskPane)
	{
		m_pTaskPane = pTaskPane;
	}

	CustomTaskPane* GetTaskpane()
	{
		return m_pTaskPane;
	}
};

