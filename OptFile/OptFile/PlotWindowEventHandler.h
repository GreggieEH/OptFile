#pragma once
class CPlotWindowEventHandler
{
public:
	CPlotWindowEventHandler();
	~CPlotWindowEventHandler();
	virtual void		OnPlotWindowMouseMove(
							double			x,
							double			y);
};

