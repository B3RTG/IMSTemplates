using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

using RabbitMQ.Client;
using RabbitMQ.Client.Events;

namespace IMSClasses.RabbitMQ
{
    public class MessageQueue
    {
        private IConnection _conn;
        private IModel _channel;
        //private BasicGetResult _lastResult;
        private QueueingBasicConsumer _consumer;
        private BasicDeliverEventArgs _eventConsumer;
        private string _queueName;
        //private string _exchange;


        //_________________________________________________________________________________________________
        // Constructor, it initializes the message queue to just consume messages
        public MessageQueue(string server, string exchange, string queue)
        {
            try
            {
                ConnectionFactory factory;

                /*server = (server == null) ? ConfigurationManager.AppSettings["queueServer"] : server;
                _queueName = (queue == null) ? ConfigurationManager.AppSettings["queueName"] : queue;*/
                _queueName = queue;

                factory = new ConnectionFactory();
                factory.UserName = "guest";
                factory.Password = "guest";
                factory.Port = 5672; // default is 5672
                factory.VirtualHost = "/"; // default is "/"
                factory.HostName = server;

                _conn = factory.CreateConnection();
                _channel = _conn.CreateModel();


                _channel.QueueDeclare(_queueName, true, false, false, null); // it will save to the disk          

                _channel.BasicQos(0, 1, false); // get just 1 message

                // We configure the class to have a event consumer
                _consumer = new QueueingBasicConsumer(_channel);
                _channel.BasicConsume(_queueName, false, _consumer);
            }
            catch (Exception e)
            {
                throw new Exception(e.Message);
            }
        }

        //_________________________________________________________________________________________________
        // 
        public int waitConsume()
        {
            int requestId;

            try
            {
                _eventConsumer = (BasicDeliverEventArgs)_consumer.Queue.Dequeue();
                string str = System.Text.Encoding.UTF8.GetString(_eventConsumer.Body);
                requestId = Int32.Parse(str);
            }
            catch (Exception e)
            {
                //throw AnalyticsError.getException(AnalyticsErrors.queueServerConsume, e);
                throw new Exception(e.Message);
            }

            return requestId;

        }

        //_________________________________________________________________________________________________
        // insert
        public void addMessage(int id)
        {
            try
            {
                IBasicProperties properties = _channel.CreateBasicProperties();
                properties.DeliveryMode = 2;//persistent
                properties.SetPersistent(true);

                _channel.BasicPublish("", _queueName, properties, Encoding.UTF8.GetBytes(id.ToString()));
                //_channel.Close();
            }
            catch (Exception e)
            {
                //throw AnalyticsError.getException(AnalyticsErrors.queueServerInsert, e);
                throw new Exception(e.Message);
            }
        }

        //_________________________________________________________________________________________________
        // Clear the queue
        public void purgeQueue()
        {
            _channel.QueuePurge(_queueName); //DEBUG for delete the queue, if the definition has changed
            while (true)
            {
                this.waitConsume();
                this.markLastMessageAsProcessed();
            }
        }

        //_________________________________________________________________________________________________
        // Mark the message as recieved
        public void markLastMessageAsProcessed()
        {
            try
            {
                // acknowledge receipt of the message
                //_channel.BasicAck(_lastResult.DeliveryTag, false);
                _channel.BasicAck(_eventConsumer.DeliveryTag, false);
            }
            catch (Exception e)
            {
                //throw AnalyticsError.getException(AnalyticsErrors.queueServerFinished, e);
                throw new Exception(e.Message);
            }
        }

        //_________________________________________________________________________________________________
        // Mark the message as not recieved
        public void markLastMessageAsNotProcessed()
        {
            try
            {
                // acknowledge receipt of the message
                _channel.BasicNack(_eventConsumer.DeliveryTag, false, true);
            }
            catch (Exception e)
            {
                //throw AnalyticsError.getException(AnalyticsErrors.queueServerFinished, e);
                throw new Exception(e.Message);
            }
        }
        //_________________________________________________________________________________________________
        // Close the message queue server connection
        public void close()
        {
            _channel.Close(200, "Bye");
            _conn.Close();
        }
    }
}
