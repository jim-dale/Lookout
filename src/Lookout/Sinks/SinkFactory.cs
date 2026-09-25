namespace Lookout.Sinks;

internal class SinkFactory
{
    internal ContactSink CreateContactSink(string[] propertyNames)
    {
        return new TableContactSink(propertyNames);
    }
}
