import json
import os
from datetime import datetime, timedelta
from aliyunsdkcore.client import AcsClient
from aliyunsdkcms.request.v20190101.DescribeMetricListRequest import DescribeMetricListRequest
import pandas as pd
import matplotlib.pyplot as plt
from docx import Document
from docx.shared import Inches
from datetime import timezone
import pytz

# Configuration - Fill these values before running
ACCESS_KEY_ID = ''
ACCESS_KEY_SECRET = ''
REGION_ID = 'cn-hongkong'

# Instance Configuration
ECS_INSTANCE_ID = 'i-j6c5wy4k8s5yc4rf9w0x'
RDS_INSTANCE_ID = 'rm-3ns01fc55d40c405n'

def get_cloud_monitor_data(client, namespace, metric_name, instance_id, start_time=None, end_time=None):
    """Fetch monitoring data from Alibaba Cloud Monitor"""
    if start_time is None or end_time is None:
        end_time = datetime.now()
        start_time = end_time - timedelta(days=14)
    
    print(f"\nRequesting data for {metric_name}:")
    print(f"Start time: {start_time}")
    print(f"End time: {end_time}")
    
    # Split the time range into 3-day chunks to ensure we get all data
    chunk_size = timedelta(days=3)
    current_start = start_time
    all_datapoints = []
    
    while current_start < end_time:
        current_end = min(current_start + chunk_size, end_time)
        
        # Format timestamps in the required format (Unix timestamp in milliseconds)
        start_timestamp = int(current_start.timestamp()) * 1000
        end_timestamp = int(current_end.timestamp()) * 1000
        
        print(f"\nFetching chunk: {current_start} to {current_end}")
        
        request = DescribeMetricListRequest()
        request.set_accept_format('json')
        request.set_Namespace(namespace)
        request.set_MetricName(metric_name)
        request.set_StartTime(start_timestamp)
        request.set_EndTime(end_timestamp)
        request.set_Dimensions(f'{{"instanceId": "{instance_id}"}}')
        request.set_Period('7200')  # 5-minute intervals for better data coverage
        
        response = client.do_action_with_exception(request)
        response_data = json.loads(response)
        
        if 'Datapoints' in response_data:
            chunk_datapoints = response_data['Datapoints']
            if isinstance(chunk_datapoints, str):
                chunk_datapoints = json.loads(chunk_datapoints)
            all_datapoints.extend(chunk_datapoints)
            print(f"Retrieved {len(chunk_datapoints)} datapoints for this chunk")
        
        current_start = current_end
    
    print(f"\nTotal datapoints retrieved: {len(all_datapoints)}")
    return {'Datapoints': all_datapoints}

def process_metrics(raw_data):    
    # Check if Datapoints exists and is not empty
    if 'Datapoints' not in raw_data or not raw_data['Datapoints']:
        print("No datapoints found in response")
        # Return empty DataFrame and default values if no data
        df = pd.DataFrame()
        stats = {'average': 0, 'max': 0, 'min': 0}
        return df, stats, pd.DataFrame()
    
    # Convert string representation of list to actual list if needed
    if isinstance(raw_data['Datapoints'], str):
        datapoints = json.loads(raw_data['Datapoints'])
    else:
        datapoints = raw_data['Datapoints']
    
    # Create DataFrame and print structure information
    df = pd.DataFrame(datapoints)
    
    # Convert timestamp and set as index
    df['timestamp'] = pd.to_datetime(df['timestamp'], unit='ms')
    df.set_index('timestamp', inplace=True)
    
    # Print data range information
    print("\nData Range Information:")
    print(f"Earliest data point: {df.index.min()}")
    print(f"Latest data point: {df.index.max()}")
    print(f"Total number of data points: {len(df)}")
    print(f"Time span: {df.index.max() - df.index.min()}")
    
    # Use 'Value' or 'Average' column for metrics (depending on what's available)
    value_column = 'Value' if 'Value' in df.columns else 'Average'
    if value_column not in df.columns:
        print(f"\nWarning: Neither 'Value' nor 'Average' column found. Available columns: {df.columns.tolist()}")
        # Return empty data if we can't find the value column
        return pd.DataFrame(), {'average': 0, 'max': 0, 'min': 0}, pd.DataFrame()
    
    # Calculate statistics using the correct column
    stats = {
        'average': df[value_column].mean(),
        'max': df[value_column].max(),
        'min': df[value_column].min()
    }
    
    # Detect anomalies (simple threshold-based)
    anomaly_threshold = 80  # 80% utilization
    anomalies = df[df[value_column] > anomaly_threshold]
    
    return df, stats, anomalies

def generate_chart(data, title, filename, days=14):
    """Generate and save line chart"""
    plt.figure(figsize=(12, 5))
    plt.plot(data.index, data['Average'], linewidth=1, label='CPU Usage')
    plt.title(f'{title} (Last {days} Days)')
    plt.ylabel('CPU Utilization (%)')
    plt.xlabel('Date/Time')
    plt.grid(True)
    
    # Format x-axis date labels
    plt.gcf().autofmt_xdate()  # Auto-format date labels
    ax = plt.gca()
    ax.xaxis.set_major_formatter(plt.matplotlib.dates.DateFormatter('%d-%m %H:%M:%S'))
    
    plt.xticks(rotation=45)
    plt.legend()
    plt.tight_layout()
    plt.savefig(filename, dpi=150)
    plt.close()

def create_word_report(report_data, output_file='cloud_report.docx'):
    """Generate Word document report"""
    doc = Document()
    
    # Report Header
    doc.add_heading('Cloud Resource Performance Report', 0)
    doc.add_paragraph(f'Report Period: {report_data["start_time"]} to {report_data["end_time"]}')
    
    # Summary Section
    doc.add_heading('Executive Summary', level=1)
    doc.add_paragraph(f"ECS CPU Utilization Range: {report_data['ecs_cpu']['min']:.2f}% to {report_data['ecs_cpu']['max']:.2f}% (Average: {report_data['ecs_cpu']['average']:.2f}%)")
    doc.add_paragraph(f"RDS CPU Utilization Range: {report_data['rds_cpu']['min']:.2f}% to {report_data['rds_cpu']['max']:.2f}% (Average: {report_data['rds_cpu']['average']:.2f}%)")
    
    # Anomalies Section
    doc.add_heading('Critical Anomalies Detected', level=1)
    if report_data['ecs_anomalies'] or report_data['rds_anomalies']:
        doc.add_paragraph('ECS Instance Anomalies:')
        for time, value in report_data['ecs_anomalies'].items():
            doc.add_paragraph(f'• {time}: {value:.1f}% CPU Utilization')
        
        doc.add_paragraph('RDS Instance Anomalies:')
        for time, value in report_data['rds_anomalies'].items():
            doc.add_paragraph(f'• {time}: {value:.1f}% CPU Utilization')
    else:
        doc.add_paragraph('No critical anomalies detected during the reporting period.')
    
    # Charts Section
    doc.add_heading('Performance Charts', level=1)
    doc.add_heading('ECS CPU Utilization', level=2)
    doc.add_picture('ecs_cpu_chart.png', width=Inches(6))
    
    doc.add_heading('RDS CPU Utilization', level=2)
    doc.add_picture('rds_cpu_chart.png', width=Inches(6))
    
    # Save document
    doc.save(output_file)
    print(f'Report generated: {output_file}')

def main():
    # Initialize Alibaba Cloud client
    client = AcsClient(ACCESS_KEY_ID, ACCESS_KEY_SECRET, REGION_ID)
    
    # Set date range for last 14 days using Hong Kong timezone (since region is Hong Kong)
    hk_tz = pytz.timezone('Asia/Hong_Kong')
    end_time = datetime.now(hk_tz)
    start_time = end_time - timedelta(days=14)
    
    # Remove timezone info as Alibaba Cloud API expects UTC
    end_time = end_time.replace(tzinfo=None)
    start_time = start_time.replace(tzinfo=None)
    
    print("\nTime Range for Data Collection:")
    print(f"Current time (HK): {datetime.now(hk_tz)}")
    print(f"Requesting from: {start_time} to {end_time}")
    
    monitoring_days = 14
    
    # Collect ECS CPU metrics
    ecs_data = get_cloud_monitor_data(
        client, 
        namespace='acs_ecs_dashboard',
        metric_name='CPUUtilization',
        instance_id=ECS_INSTANCE_ID,
        start_time=start_time,
        end_time=end_time
    )
    
    # Collect RDS CPU metrics
    rds_data = get_cloud_monitor_data(
        client,
        namespace='acs_rds_dashboard',
        metric_name='CpuUsage',
        instance_id=RDS_INSTANCE_ID,
        start_time=start_time,
        end_time=end_time
    )
    
    # Process data
    ecs_df, ecs_stats, ecs_anomalies = process_metrics(ecs_data)
    rds_df, rds_stats, rds_anomalies = process_metrics(rds_data)
    
    # Generate charts
    generate_chart(ecs_df, 'ECS CPU Utilization', 'ecs_cpu_chart.png', days=monitoring_days)
    generate_chart(rds_df, 'RDS CPU Utilization', 'rds_cpu_chart.png', days=monitoring_days)
    
    # Prepare report data
    report_data = {
        'start_time': start_time.strftime('%Y-%m-%d %H:%M:%S'),
        'end_time': end_time.strftime('%Y-%m-%d %H:%M:%S'),
        'ecs_cpu': ecs_stats,
        'rds_cpu': rds_stats,
        'ecs_anomalies': ecs_anomalies.iloc[:, 0].to_dict() if not ecs_anomalies.empty else {},
        'rds_anomalies': rds_anomalies.iloc[:, 0].to_dict() if not rds_anomalies.empty else {}
    }
    
    # Generate Word document
    create_word_report(report_data)
    
    # Cleanup temporary files
    os.remove('ecs_cpu_chart.png')
    os.remove('rds_cpu_chart.png')

if __name__ == "__main__":
    main()