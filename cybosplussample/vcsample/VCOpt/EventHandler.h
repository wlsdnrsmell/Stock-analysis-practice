// CEventHandler.h

#ifndef __EVNET_HANDLER_H__
#define __EVNET_HANDLER_H__

class IEventHandler
{
public:
	virtual void Received() = 0;
};


class CEventHandler : public IDispEventImpl<
	0,
	CEventHandler,
	&DIID__IDibEvents,
	&LIBID_DSCBO1Lib,
	1,
	0>
{
public:
	void SetIEventHandler(IEventHandler* pIEventHandler) { m_pIEventHandler = pIEventHandler; }

	void __stdcall Received();

	BEGIN_SINK_MAP(CEventHandler)
		SINK_ENTRY_EX(0, DIID__IDibEvents, 1, Received)
	END_SINK_MAP()

protected:
	IEventHandler* m_pIEventHandler;
};

#endif /* __EVNET_HANDLER_H__ */
