// EventHandler.cpp

#include "stdafx.h"
#include "EventHandler.h"

void __stdcall CEventHandler::Received()
{
	ASSERT(NULL != m_pIEventHandler);

	m_pIEventHandler->Received();
}
