#!/usr/bin/env python3
import requests
import argparse
import re
from urllib.parse import urljoin, urlparse
from bs4 import BeautifulSoup
import socket
import ssl
import json
import time
from concurrent.futures import ThreadPoolExecutor, as_completed

class VulnerabilityScanner:
    def __init__(self, target_url, max_threads=10):
        self.target_url = target_url.rstrip('/')
        self.session = requests.Session()
        self.session.headers.update({
            'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/91.0.4472.124 Safari/537.36'
        })
        self.links_to_ignore = set()
        self.vulnerabilities = []
        self.max_threads = max_threads
        self.cookies = {}
        self.auth_creds = None
        self.forms = []
        self.technologies = []
        self.scan_start_time = time.time()

    def scan(self):
        print(f"\n[+] Starting scan for: {self.target_url}")
        print(f"[+] Scan started at: {time.strftime('%Y-%m-%d %H:%M:%S')}\n")
        
        # Initial discovery
        self.discover_technologies()
        self.crawl()
        
        # Vulnerability tests
        tests = [
            self.test_sql_injection,
            self.test_xss,
            self.test_directory_traversal,
            self.test_sensitive_files,
            self.test_command_injection,
            self.test_ssrf,
            self.test_idor,
            self.test_csrf,
            self.test_http_methods,
            self.test_cors,
            self.test_security_headers,
            self.test_jwt_issues,
            self.test_clickjacking,
            self.test_server_info_disclosure
        ]
        
        # Run tests with thread pooling
        with ThreadPoolExecutor(max_workers=self.max_threads) as executor:
            futures = [executor.submit(test) for test in tests]
            for future in as_completed(futures):
                try:
                    future.result()
                except Exception as e:
                    print(f"[-] Error in test: {e}")
        
        self.report_results()

    def discover_technologies(self):
        print("[+] Discovering technologies...")
        try:
            headers = self.session.get(self.target_url).headers
            
            # Check server headers
            if 'Server' in headers:
                self.technologies.append(f"Server: {headers['Server']}")
            
            # Check X-Powered-By
            if 'X-Powered-By' in headers:
                self.technologies.append(f"Powered By: {headers['X-Powered-By']}")
            
            # Check cookies for technology hints
            for cookie in self.session.cookies:
                if 'php' in cookie.name.lower():
                    self.technologies.append("PHP detected from cookies")
                    break
            
            # Check common paths
            tech_paths = {
                'wp-admin': 'WordPress',
                'joomla': 'Joomla',
                'drupal': 'Drupal',
                'laravel': 'Laravel'
            }
            
            for path, tech in tech_paths.items():
                if requests.get(f"{self.target_url}/{path}", timeout=5).status_code == 200:
                    self.technologies.append(tech)
            
            # Check for common JavaScript frameworks
            response = self.session.get(self.target_url)
            if 'react' in response.text.lower():
                self.technologies.append("ReactJS")
            if 'angular' in response.text.lower():
                self.technologies.append("Angular")
            if 'vue' in response.text.lower():
                self.technologies.append("Vue.js")
            
        except Exception as e:
            print(f"[-] Error during technology discovery: {e}")

    def crawl(self):
        print("[+] Crawling the website for links and forms...")
        try:
            response = self.session.get(self.target_url)
            soup = BeautifulSoup(response.content, "html.parser")
            
            # Extract all links
            for link in soup.find_all("a"):
                href = link.get("href")
                if href and not href.startswith(('javascript:', 'mailto:', 'tel:')):
                    full_url = urljoin(self.target_url, href)
                    if self.target_url in full_url and full_url not in self.links_to_ignore:
                        self.links_to_ignore.add(full_url)
                        print(f"[+] Found link: {full_url}")
                        # Recursive crawling (limited depth in real implementation)
            
            # Extract all forms
            for form in soup.find_all("form"):
                form_details = {
                    'action': urljoin(self.target_url, form.get('action', '')),
                    'method': form.get('method', 'get').upper(),
                    'inputs': []
                }
                
                for input_tag in form.find_all("input"):
                    form_details['inputs'].append({
                        'name': input_tag.get('name'),
                        'type': input_tag.get('type', 'text'),
                        'value': input_tag.get('value', '')
                    })
                
                self.forms.append(form_details)
                print(f"[+] Found form: {form_details['action']} ({form_details['method']})")
            
            # Extract API endpoints from JavaScript
            script_tags = soup.find_all("script")
            for script in script_tags:
                if script.get("src"):
                    continue
                api_patterns = re.findall(r'fetch\(["\'](.*?)["\']\)', script.text)
                for api in api_patterns:
                    full_api = urljoin(self.target_url, api)
                    if full_api not in self.links_to_ignore:
                        self.links_to_ignore.add(full_api)
                        print(f"[+] Found API endpoint: {full_api}")
            
        except Exception as e:
            print(f"[-] Error during crawling: {e}")

    def test_sql_injection(self):
        print("[+] Testing for SQL Injection vulnerabilities...")
        test_payloads = [
            "'", "\"", 
            "' OR '1'='1", "\" OR \"1\"=\"1", 
            "' OR 1=1--", "admin'--",
            "1 AND 1=1", "1 AND 1=2",
            "1' ORDER BY 1--", "1' ORDER BY 10--",
            "1' UNION SELECT null,username,password FROM users--"
        ]
        
        self._test_url_params(test_payloads, "SQL Injection", 
                             lambda r: any(e in r.text.lower() for e in ['error', 'syntax', 'mysql', 'ora-', 'sql']))
        
        # Test forms for SQLi
        for form in self.forms:
            for payload in test_payloads:
                data = {}
                for input_field in form['inputs']:
                    if input_field['type'] in ['text', 'password', 'hidden']:
                        data[input_field['name']] = payload
                    else:
                        data[input_field['name']] = input_field['value']
                
                try:
                    if form['method'] == 'POST':
                        response = self.session.post(form['action'], data=data)
                    else:
                        response = self.session.get(form['action'], params=data)
                    
                    if any(e in response.text.lower() for e in ['error', 'syntax', 'mysql', 'ora-', 'sql']):
                        self.vulnerabilities.append((
                            "SQL Injection", 
                            form['action'], 
                            f"Form parameter with payload: {payload}",
                            response.url
                        ))
                        print(f"[!] Potential SQL Injection in form at: {form['action']}")
                        break
                except:
                    continue

    def test_xss(self):
        print("[+] Testing for XSS vulnerabilities...")
        test_payloads = [
            "<script>alert('XSS')</script>",
            "<img src=x onerror=alert('XSS')>",
            "\"><script>alert('XSS')</script>",
            "javascript:alert('XSS')",
            "onmouseover=alert('XSS')",
            "<svg/onload=alert('XSS')>"
        ]
        
        self._test_url_params(test_payloads, "Cross-Site Scripting (XSS)", 
                            lambda r: any(p in r.text for p in test_payloads))
        
        # Test forms for XSS
        for form in self.forms:
            for payload in test_payloads:
                data = {}
                for input_field in form['inputs']:
                    if input_field['type'] in ['text', 'textarea', 'search']:
                        data[input_field['name']] = payload
                    else:
                        data[input_field['name']] = input_field['value']
                
                try:
                    if form['method'] == 'POST':
                        response = self.session.post(form['action'], data=data)
                    else:
                        response = self.session.get(form['action'], params=data)
                    
                    if payload in response.text:
                        self.vulnerabilities.append((
                            "Cross-Site Scripting (XSS)", 
                            form['action'], 
                            f"Form parameter with payload: {payload}",
                            response.url
                        ))
                        print(f"[!] Potential XSS in form at: {form['action']}")
                        break
                except:
                    continue

    def test_directory_traversal(self):
        print("[+] Testing for Directory Traversal vulnerabilities...")
        test_payloads = [
            "../../../../etc/passwd",
            "../../../../windows/win.ini",
            "%2e%2e%2f%2e%2e%2f%2e%2e%2fetc%2fpasswd",
            "..%2F..%2F..%2Fetc%2Fpasswd",
            "....//....//etc/passwd"
        ]
        
        base_url = self.target_url + "/" if not self.target_url.endswith("/") else self.target_url
        
        for payload in test_payloads:
            test_url = base_url + payload
            try:
                response = self.session.get(test_url)
                if "root:" in response.text or "[extensions]" in response.text:
                    self.vulnerabilities.append((
                        "Directory Traversal", 
                        test_url, 
                        payload,
                        response.url
                    ))
                    print(f"[!] Potential Directory Traversal at: {test_url}")
            except:
                continue

    def test_sensitive_files(self):
        print("[+] Testing for sensitive files...")
        sensitive_files = [
            "robots.txt", ".git/", ".env", "config.php", "wp-config.php",
            "phpinfo.php", "admin/", "backup/", "phpmyadmin/", ".htaccess",
            ".DS_Store", "web.config", "server-status", "debug.log",
            "package.json", "composer.json", "yarn.lock", "Gemfile.lock",
            "config.yml", "settings.py", "secrets.yml", "credentials.json"
        ]
        
        base_url = self.target_url + "/" if not self.target_url.endswith("/") else self.target_url
        
        for file in sensitive_files:
            test_url = base_url + file
            try:
                response = self.session.get(test_url)
                if response.status_code == 200:
                    content_type = response.headers.get('Content-Type', '')
                    if 'text/html' not in content_type or file.endswith(('.php', '.json', '.yml', '.py')):
                        self.vulnerabilities.append((
                            "Sensitive File Exposure", 
                            test_url, 
                            "",
                            response.url
                        ))
                        print(f"[!] Sensitive file found: {test_url}")
            except:
                continue

    def test_command_injection(self):
        print("[+] Testing for Command Injection vulnerabilities...")
        test_payloads = [
            ";id", "|id", "`id`", "$(id)", "|| id", "&& id",
            "; whoami", "| whoami", "`whoami`", "$(whoami)"
        ]
        
        self._test_url_params(test_payloads, "Command Injection", 
                            lambda r: any(e in r.text.lower() for e in ['uid=', 'gid=', 'groups=', 'windows nt']))

    def test_ssrf(self):
        print("[+] Testing for Server-Side Request Forgery (SSRF)...")
        test_payloads = [
            "http://169.254.169.254/latest/meta-data/",
            "http://localhost/admin",
            "http://127.0.0.1:8080"
        ]
        
        self._test_url_params(test_payloads, "SSRF Potential", 
                            lambda r: any(e in r.text.lower() for e in ['aws', 'ec2', 'metadata', 'admin']))

    def test_idor(self):
        print("[+] Testing for Insecure Direct Object References (IDOR)...")
        # This would need specific test cases based on application context
        pass

    def test_csrf(self):
        print("[+] Testing for CSRF vulnerabilities...")
        for form in self.forms:
            if not any(input_field.get('name', '').lower() in ['csrf', 'csrftoken', '_token'] 
               for input_field in form['inputs']):
                self.vulnerabilities.append((
                    "Potential CSRF (Missing Token)", 
                    form['action'], 
                    "Form missing CSRF token",
                    form['action']
                ))
                print(f"[!] Potential CSRF vulnerability in form at: {form['action']}")

    def test_http_methods(self):
        print("[+] Testing for insecure HTTP methods...")
        test_urls = [self.target_url] + list(self.links_to_ignore)[:10]  # Limit to first 10 URLs
        methods = ['OPTIONS', 'TRACE', 'PUT', 'DELETE']
        
        for url in test_urls:
            for method in methods:
                try:
                    response = self.session.request(method, url)
                    if response.status_code not in [405, 501]:
                        self.vulnerabilities.append((
                            "Insecure HTTP Method Allowed", 
                            url, 
                            f"Method: {method}",
                            url
                        ))
                        print(f"[!] {method} method allowed at: {url}")
                except:
                    continue

    def test_cors(self):
        print("[+] Testing for CORS misconfigurations...")
        test_headers = {
            'Origin': 'https://evil.com',
            'Access-Control-Request-Method': 'GET'
        }
        
        try:
            response = self.session.get(self.target_url, headers=test_headers)
            cors_headers = response.headers.get('Access-Control-Allow-Origin', '')
            if cors_headers == '*' or 'evil.com' in cors_headers:
                self.vulnerabilities.append((
                    "CORS Misconfiguration", 
                    self.target_url, 
                    f"Access-Control-Allow-Origin: {cors_headers}",
                    self.target_url
                ))
                print(f"[!] Insecure CORS configuration at: {self.target_url}")
        except:
            pass

    def test_security_headers(self):
        print("[+] Testing for missing security headers...")
        required_headers = [
            'X-Content-Type-Options',
            'X-Frame-Options',
            'Content-Security-Policy',
            'X-XSS-Protection',
            'Strict-Transport-Security'
        ]
        
        try:
            response = self.session.get(self.target_url)
            missing_headers = [h for h in required_headers if h not in response.headers]
            
            if missing_headers:
                self.vulnerabilities.append((
                    "Missing Security Headers", 
                    self.target_url, 
                    f"Missing headers: {', '.join(missing_headers)}",
                    self.target_url
                ))
                print(f"[!] Missing security headers at: {self.target_url}")
        except:
            pass

    def test_jwt_issues(self):
        print("[+] Testing for JWT issues...")
        # This would need specific JWT testing logic
        pass

    def test_clickjacking(self):
        print("[+] Testing for Clickjacking vulnerabilities...")
        try:
            response = self.session.get(self.target_url)
            if 'X-Frame-Options' not in response.headers:
                self.vulnerabilities.append((
                    "Clickjacking Potential", 
                    self.target_url, 
                    "Missing X-Frame-Options header",
                    self.target_url
                ))
                print(f"[!] Potential Clickjacking vulnerability at: {self.target_url}")
        except:
            pass

    def test_server_info_disclosure(self):
        print("[+] Testing for server information disclosure...")
        try:
            response = self.session.get(self.target_url)
            headers = response.headers
            
            info_disclosed = []
            if 'Server' in headers:
                info_disclosed.append(f"Server: {headers['Server']}")
            if 'X-Powered-By' in headers:
                info_disclosed.append(f"Powered By: {headers['X-Powered-By']}")
            if 'X-AspNet-Version' in headers:
                info_disclosed.append(f"ASP.NET Version: {headers['X-AspNet-Version']}")
            
            if info_disclosed:
                self.vulnerabilities.append((
                    "Information Disclosure", 
                    self.target_url, 
                    f"Disclosed: {', '.join(info_disclosed)}",
                    self.target_url
                ))
                print(f"[!] Server information disclosure at: {self.target_url}")
        except:
            pass

    def _test_url_params(self, payloads, vuln_type, detection_logic):
        for url in self.links_to_ignore:
            if "=" in url:  # Only test URLs with parameters
                for payload in payloads:
                    test_url = url.replace("=", f"={payload}")
                    try:
                        response = self.session.get(test_url)
                        if detection_logic(response):
                            self.vulnerabilities.append((
                                vuln_type, 
                                url, 
                                payload,
                                response.url
                            ))
                            print(f"[!] Potential {vuln_type} at: {url}")
                            break
                    except:
                        continue

    def report_results(self):
        scan_duration = time.time() - self.scan_start_time
        print("\n[+] Scan completed!")
        print(f"[+] Scan duration: {scan_duration:.2f} seconds")
        print(f"[+] Technologies detected: {', '.join(self.technologies) if self.technologies else 'None detected'}")
        print(f"[+] Found {len(self.vulnerabilities)} potential vulnerabilities:\n")
        
        if not self.vulnerabilities:
            print("No vulnerabilities found!")
            return
        
        # Group vulnerabilities by type
        vuln_by_type = {}
        for vuln in self.vulnerabilities:
            if vuln[0] not in vuln_by_type:
                vuln_by_type[vuln[0]] = []
            vuln_by_type[vuln[0]].append(vuln)
        
        # Print results
        for vuln_type, vulns in vuln_by_type.items():
            print(f"\n=== {vuln_type} ({len(vulns)} found) ===")
            for i, vuln in enumerate(vulns, 1):
                print(f"\n{i}. URL: {vuln[1]}")
                if vuln[2]:
                    print(f"   Payload/Details: {vuln[2]}")
                print(f"   Tested URL: {vuln[3]}")
            print("-" * 60)
        
        print("\n[+] Scan finished at: " + time.strftime('%Y-%m-%d %H:%M:%S'))

if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="Enhanced Web Application Vulnerability Scanner")
    parser.add_argument("-u", "--url", required=True, help="Target URL to scan")
    parser.add_argument("-t", "--threads", type=int, default=10, help="Maximum threads to use")
    parser.add_argument("-a", "--auth", help="Authentication credentials (user:pass)")
    args = parser.parse_args()
    
    scanner = VulnerabilityScanner(args.url, args.threads)
    
    if args.auth:
        try:
            username, password = args.auth.split(":")
            scanner.auth_creds = (username, password)
            scanner.session.auth = (username, password)
        except:
            print("[-] Invalid authentication format. Use user:pass")
            exit(1)
    
    scanner.scan()
