function y = PIECEWISELINE(x,a,b,c,d,e)
% PIECEWISELINE   A line made of two pieces
% that is not continuous.

y=x;
for i=1:length(x)
    if x(i)<e
        y(i)=a+b*x(i);
    else
        y(i)=c+d*x(i);
    end

end